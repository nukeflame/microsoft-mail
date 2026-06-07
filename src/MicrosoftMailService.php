<?php

namespace Nukeflame\MicrosoftMail;

use App\Models\UserMailToken;
use Carbon\Carbon;
use GuzzleHttp\Client;

class MicrosoftMailService
{
    private const GRAPH_URL  = 'https://graph.microsoft.com/v1.0';
    private const TOKEN_BASE = 'https://login.microsoftonline.com/%s/oauth2/v2.0/token';

    private const FOLDER_MAP = [
        'inbox'   => 'inbox',
        'sent'    => 'sentitems',
        'drafts'  => 'drafts',
        'archive' => 'archive',
        'junk'    => 'junkemail',
        'deleted' => 'deleteditems',
    ];

    private const MESSAGE_SELECT = 'id,subject,from,bodyPreview,receivedDateTime,isRead,flag,hasAttachments,categories,parentFolderId';
    private const MESSAGE_DETAIL_SELECT = 'id,subject,from,toRecipients,ccRecipients,body,bodyPreview,receivedDateTime,isRead,flag,hasAttachments,categories,parentFolderId';

    public function __construct(
        private readonly MailRepository $mailRepository,
        private readonly Client $httpClient,
    ) {}

    public function getFolders(int $userId): array
    {
        $token = $this->resolveToken($userId);

        $resp = $this->graphGet($token, '/me/mailFolders', [
            '$top' => 20,
        ]);

        return $resp['value'] ?? [];
    }

    public function getUnreadCounts(int $userId): array
    {
        $token   = $this->resolveToken($userId);
        $folders = $this->getFolders($userId);
        $counts  = [];

        foreach ($folders as $folder) {
            $name = $folder['wellKnownName'] ?? null;
            if (! $name) {
                continue;
            }
            $localFolder = array_search($name, self::FOLDER_MAP, true);
            if ($localFolder !== false) {
                $counts[$localFolder] = $folder['unreadItemCount'] ?? 0;
            }
        }

        $starred = $this->graphGet($token, '/me/messages', [
            '$filter' => "flag/flagStatus eq 'flagged' and isRead eq false",
            '$count'  => 'true',
            '$top'    => 1,
            '$select' => 'id',
        ]);
        $counts['starred'] = $starred['@odata.count'] ?? 0;

        return $counts;
    }

    public function getMessages(int $userId, string $folder, ?string $query = null, int $top = 50, int $skip = 0): array
    {
        $token = $this->resolveToken($userId);

        $params = [
            '$top'    => $top,
            '$skip'   => $skip,
            '$select' => self::MESSAGE_SELECT,
        ];

        if ($folder === 'starred') {
            $endpoint           = '/me/messages';
            $params['$filter']  = "flag/flagStatus eq 'flagged'";
            $params['$orderby'] = 'receivedDateTime desc';
        } else {
            $graphFolder        = self::FOLDER_MAP[$folder] ?? $folder;
            $endpoint           = "/me/mailFolders/{$graphFolder}/messages";
            $params['$orderby'] = 'receivedDateTime desc';
        }

        if ($query) {
            $params['$search'] = '"' . str_replace('"', '', $query) . '"';
            unset($params['$orderby']);
        }

        $resp = $this->graphGet($token, $endpoint, $params);

        return [
            'messages' => array_map(fn($m) => $this->mapMessage($m, $folder), $resp['value'] ?? []),
            'nextLink' => $resp['@odata.nextLink'] ?? null,
        ];
    }

    public function getMessage(int $userId, string $messageId): array
    {
        $token = $this->resolveToken($userId);
        $msg   = $this->graphGet($token, "/me/messages/{$messageId}", [
            '$select' => self::MESSAGE_DETAIL_SELECT,
        ]);

        return $this->mapMessage($msg, $this->resolveFolderFromParent($msg['parentFolderId'] ?? ''));
    }

    public function markRead(int $userId, string $messageId, bool $read): array
    {
        $token = $this->resolveToken($userId);
        $msg   = $this->graphPatch($token, "/me/messages/{$messageId}", ['isRead' => $read]);

        return $this->mapMessage($msg, 'inbox');
    }

    public function toggleFlag(int $userId, string $messageId): array
    {
        $token   = $this->resolveToken($userId);
        $current = $this->graphGet($token, "/me/messages/{$messageId}", ['$select' => 'id,flag']);

        $isFlagged = ($current['flag']['flagStatus'] ?? '') === 'flagged';
        $newStatus = $isFlagged ? 'notFlagged' : 'flagged';

        $msg = $this->graphPatch($token, "/me/messages/{$messageId}", [
            'flag' => ['flagStatus' => $newStatus],
        ]);

        return $this->mapMessage($msg, 'inbox');
    }

    public function moveMessage(int $userId, string $messageId, string $targetFolder): array
    {
        $token       = $this->resolveToken($userId);
        $graphFolder = self::FOLDER_MAP[$targetFolder] ?? $targetFolder;

        $msg = $this->graphPost($token, "/me/messages/{$messageId}/move", [
            'destinationId' => $graphFolder,
        ]);

        return $this->mapMessage($msg, $targetFolder);
    }

    public function deleteMessage(int $userId, string $messageId): void
    {
        $token = $this->resolveToken($userId);
        $this->graphDelete($token, "/me/messages/{$messageId}");
    }

    public function sendMail(int $userId, array $data): void
    {
        $token = $this->resolveToken($userId);

        $payload = [
            'message' => [
                'subject'      => $data['subject'] ?? '(no subject)',
                'body'         => ['contentType' => 'HTML', 'content' => $data['body'] ?? ''],
                'toRecipients' => $this->parseRecipients($data['to'] ?? ''),
            ],
            'saveToSentItems' => true,
        ];

        if (! empty($data['cc'])) {
            $payload['message']['ccRecipients'] = $this->parseRecipients($data['cc']);
        }

        if (! empty($data['bcc'])) {
            $payload['message']['bccRecipients'] = $this->parseRecipients($data['bcc']);
        }

        if (! empty($data['attachments'])) {
            $payload['message']['attachments'] = array_map(
                static fn (array $attachment): array => [
                    '@odata.type'  => '#microsoft.graph.fileAttachment',
                    'name'         => $attachment['name'],
                    'contentType'  => $attachment['contentType'] ?? 'application/octet-stream',
                    'contentBytes' => $attachment['contentBytes'],
                ],
                $data['attachments'],
            );
        }

        $this->graphPost($token, '/me/sendMail', $payload);
    }

    public function reply(int $userId, string $messageId, string $body, bool $replyAll = false): void
    {
        $token    = $this->resolveToken($userId);
        $endpoint = $replyAll
            ? "/me/messages/{$messageId}/replyAll"
            : "/me/messages/{$messageId}/reply";

        $this->graphPost($token, $endpoint, [
            'message' => ['body' => ['contentType' => 'HTML', 'content' => $body]],
            'comment' => $body,
        ]);
    }

    public function forward(int $userId, string $messageId, string $to, string $comment = ''): void
    {
        $token = $this->resolveToken($userId);

        $this->graphPost($token, "/me/messages/{$messageId}/forward", [
            'toRecipients' => $this->parseRecipients($to),
            'comment'      => $comment,
        ]);
    }

    public function saveDraft(int $userId, array $data): array
    {
        $token = $this->resolveToken($userId);

        $msg = $this->graphPost($token, '/me/messages', [
            'subject'      => $data['subject'] ?? '(no subject)',
            'body'         => ['contentType' => 'HTML', 'content' => $data['body'] ?? ''],
            'toRecipients' => $this->parseRecipients($data['to'] ?? ''),
        ]);

        return $this->mapMessage($msg, 'drafts');
    }

    public function updateDraft(int $userId, string $messageId, array $data): array
    {
        $token = $this->resolveToken($userId);

        $msg = $this->graphPatch($token, "/me/messages/{$messageId}", [
            'subject'      => $data['subject'] ?? '(no subject)',
            'body'         => ['contentType' => 'HTML', 'content' => $data['body'] ?? ''],
            'toRecipients' => $this->parseRecipients($data['to'] ?? ''),
        ]);

        return $this->mapMessage($msg, 'drafts');
    }

    public function sendDraft(int $userId, string $messageId): void
    {
        $token = $this->resolveToken($userId);
        $this->graphPost($token, "/me/messages/{$messageId}/send", []);
    }

    private function resolveToken(int $userId): string
    {
        $record = $this->mailRepository->findByUser($userId);

        if (! $record) {
            throw new \RuntimeException('Microsoft mailbox not connected. Please connect your account first.', 403);
        }

        if ($record->isExpired()) {
            $record = $this->refreshAccessToken($record);
        }

        return $record->access_token;
    }

    private function refreshAccessToken(UserMailToken $record): UserMailToken
    {
        $tokenUrl = sprintf(self::TOKEN_BASE, config('services.microsoft_mail.tenant'));

        $resp = $this->httpClient->post($tokenUrl, [
            'form_params' => [
                'client_id'     => config('services.microsoft_mail.client_id'),
                'client_secret' => config('services.microsoft_mail.client_secret'),
                'grant_type'    => 'refresh_token',
                'refresh_token' => $record->refresh_token,
                'scope'         => config('services.microsoft_mail.scopes'),
            ],
        ]);

        $data = json_decode($resp->getBody()->getContents(), true);

        $this->mailRepository->update($record, [
            'access_token'  => $data['access_token'],
            'refresh_token' => $data['refresh_token'] ?? $record->refresh_token,
            'expires_at'    => now()->addSeconds((int) $data['expires_in'] - 60),
        ]);

        return $record->fresh();
    }

    private function graphGet(string $token, string $path, array $params = []): array
    {
        $url = self::GRAPH_URL . $path;
        if ($params) {
            $url .= '?' . http_build_query($params);
        }

        $resp = $this->httpClient->get($url, ['headers' => $this->authHeaders($token)]);

        return json_decode($resp->getBody()->getContents(), true) ?? [];
    }

    private function graphPost(string $token, string $path, array $body): array
    {
        $resp = $this->httpClient->post(self::GRAPH_URL . $path, [
            'headers' => array_merge($this->authHeaders($token), ['Content-Type' => 'application/json']),
            'json'    => $body,
        ]);

        $contents = $resp->getBody()->getContents();

        return $contents ? (json_decode($contents, true) ?? []) : [];
    }

    private function graphPatch(string $token, string $path, array $body): array
    {
        $resp = $this->httpClient->patch(self::GRAPH_URL . $path, [
            'headers' => array_merge($this->authHeaders($token), ['Content-Type' => 'application/json']),
            'json'    => $body,
        ]);

        return json_decode($resp->getBody()->getContents(), true) ?? [];
    }

    private function graphDelete(string $token, string $path): void
    {
        $this->httpClient->delete(self::GRAPH_URL . $path, ['headers' => $this->authHeaders($token)]);
    }

    private function authHeaders(string $token): array
    {
        return [
            'Authorization' => "Bearer {$token}",
            'Accept'        => 'application/json',
        ];
    }

    private function mapMessage(array $msg, string $folder): array
    {
        $flagStatus = $msg['flag']['flagStatus'] ?? 'notFlagged';

        return [
            'id'            => $msg['id'],
            'folder'        => $folder,
            'from'          => $msg['from']['emailAddress']['name'] ?? 'Unknown',
            'address'       => $msg['from']['emailAddress']['address'] ?? '',
            'subject'       => $msg['subject'] ?? '(no subject)',
            'preview'       => $msg['bodyPreview'] ?? '',
            'body'          => $msg['body']['content'] ?? ($msg['bodyPreview'] ?? ''),
            'time'          => $this->formatTime($msg['receivedDateTime'] ?? now()->toIso8601String()),
            'read'          => $msg['isRead'] ?? false,
            'starred'       => $flagStatus === 'flagged',
            'flagged'       => $flagStatus === 'flagged',
            'hasAttachment' => $msg['hasAttachments'] ?? false,
            'category'      => $msg['categories'][0] ?? null,
        ];
    }

    private function formatTime(string $datetime): string
    {
        $dt  = Carbon::parse($datetime)->setTimezone(config('app.timezone', 'UTC'));
        $now = Carbon::now();

        if ($dt->isToday()) {
            return $dt->format('H:i');
        }

        if ($dt->isYesterday()) {
            return 'Yesterday';
        }

        if ($dt->diffInDays($now) < 7) {
            return $dt->format('D');
        }

        return $dt->format('d M');
    }

    private function parseRecipients(string $addresses): array
    {
        return array_values(array_filter(array_map(function (string $addr): ?array {
            $addr = trim($addr);
            if (! $addr) {
                return null;
            }

            if (preg_match('/^(.+?)\s*<(.+?)>$/', $addr, $m)) {
                return ['emailAddress' => ['name' => trim($m[1]), 'address' => trim($m[2])]];
            }

            return ['emailAddress' => ['name' => $addr, 'address' => $addr]];
        }, explode(',', $addresses))));
    }

    private function resolveFolderFromParent(string $parentFolderId): string
    {
        return 'inbox';
    }
}
