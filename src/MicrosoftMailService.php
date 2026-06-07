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

    /**
     * Attachments at or below this size can be inlined as base64 in a single sendMail/attachments
     * call. Microsoft Graph rejects larger file attachments on those endpoints, requiring an
     * upload session instead.
     *
     * @see https://learn.microsoft.com/en-us/graph/outlook-large-attachments
     */
    private const INLINE_ATTACHMENT_LIMIT = 3 * 1024 * 1024;

    /** Upload session chunk size — must be a multiple of 320 KiB per Graph's guidance. */
    private const UPLOAD_CHUNK_SIZE = 320 * 1024 * 12;

    /**
     * @param  array<string, mixed>  $data  Expects 'attachments' as a list of
     *                                      {name, contentType, contents} where 'contents' is the
     *                                      raw (non-base64) file bytes.
     */
    public function sendMail(int $userId, array $data): void
    {
        $token       = $this->resolveToken($userId);
        $attachments = $this->normalizeAttachments($data['attachments'] ?? []);

        $hasLargeAttachment = false;
        foreach ($attachments as $attachment) {
            if ($attachment['size'] > self::INLINE_ATTACHMENT_LIMIT) {
                $hasLargeAttachment = true;
                break;
            }
        }

        if ($hasLargeAttachment) {
            $this->sendViaDraftWithAttachments($token, $data, $attachments);

            return;
        }

        $message = $this->buildMessagePayload($data);

        if ($attachments !== []) {
            $message['attachments'] = array_map(
                fn (array $attachment): array => $this->fileAttachmentPayload($attachment),
                $attachments,
            );
        }

        $this->graphPost($token, '/me/sendMail', [
            'message'         => $message,
            'saveToSentItems' => true,
        ]);
    }

    /**
     * Sends a message that has at least one attachment too large to inline: create a draft,
     * attach files individually (large ones via a chunked upload session), then send the draft.
     *
     * @param  array<int, array{name: string, contentType: string, contents: string, size: int}>  $attachments
     */
    private function sendViaDraftWithAttachments(string $token, array $data, array $attachments): void
    {
        $draft     = $this->graphPost($token, '/me/messages', $this->buildMessagePayload($data));
        $messageId = $draft['id'] ?? null;

        if (! $messageId) {
            throw new \RuntimeException('Microsoft Graph did not return a draft message id while preparing attachments.');
        }

        foreach ($attachments as $attachment) {
            if ($attachment['size'] > self::INLINE_ATTACHMENT_LIMIT) {
                $this->uploadLargeAttachment($token, $messageId, $attachment);
            } else {
                $this->graphPost($token, "/me/messages/{$messageId}/attachments", $this->fileAttachmentPayload($attachment));
            }
        }

        $this->graphPost($token, "/me/messages/{$messageId}/send", []);
    }

    /**
     * Uploads an attachment that exceeds Graph's inline limit using a resumable upload session,
     * streaming the raw bytes in fixed-size chunks via the Content-Range header.
     *
     * @param  array{name: string, contentType: string, contents: string, size: int}  $attachment
     */
    private function uploadLargeAttachment(string $token, string $messageId, array $attachment): void
    {
        $session = $this->graphPost($token, "/me/messages/{$messageId}/attachments/createUploadSession", [
            'AttachmentItem' => [
                'attachmentType' => 'file',
                'name'           => $attachment['name'],
                'contentType'    => $attachment['contentType'],
                'size'           => $attachment['size'],
            ],
        ]);

        $uploadUrl = $session['uploadUrl'] ?? null;

        if (! $uploadUrl) {
            throw new \RuntimeException("Microsoft Graph did not return an upload session for attachment \"{$attachment['name']}\".");
        }

        $size   = $attachment['size'];
        $offset = 0;

        while ($offset < $size) {
            $length = min(self::UPLOAD_CHUNK_SIZE, $size - $offset);

            // The upload URL is pre-authenticated; sending a Bearer token alongside it is rejected.
            $this->httpClient->put($uploadUrl, [
                'headers' => [
                    'Content-Length' => (string) $length,
                    'Content-Range'  => sprintf('bytes %d-%d/%d', $offset, $offset + $length - 1, $size),
                ],
                'body' => substr($attachment['contents'], $offset, $length),
            ]);

            $offset += $length;
        }
    }

    /**
     * @param  array<int, array{name?: string, contentType?: string, contents?: string}>  $attachments
     * @return array<int, array{name: string, contentType: string, contents: string, size: int}>
     */
    private function normalizeAttachments(array $attachments): array
    {
        return array_map(static function (array $attachment): array {
            $contents = $attachment['contents'] ?? '';

            return [
                'name'        => $attachment['name'] ?? 'attachment',
                'contentType' => $attachment['contentType'] ?? 'application/octet-stream',
                'contents'    => $contents,
                'size'        => strlen($contents),
            ];
        }, $attachments);
    }

    /**
     * @param  array{name: string, contentType: string, contents: string, size: int}  $attachment
     * @return array<string, string>
     */
    private function fileAttachmentPayload(array $attachment): array
    {
        return [
            '@odata.type'  => '#microsoft.graph.fileAttachment',
            'name'         => $attachment['name'],
            'contentType'  => $attachment['contentType'],
            'contentBytes' => base64_encode($attachment['contents']),
        ];
    }

    /**
     * @return array<string, mixed>
     */
    private function buildMessagePayload(array $data): array
    {
        $message = [
            'subject'      => $data['subject'] ?? '(no subject)',
            'body'         => ['contentType' => 'HTML', 'content' => $this->htmlFromPlainText($data['body'] ?? '')],
            'toRecipients' => $this->parseRecipients($data['to'] ?? ''),
        ];

        if (! empty($data['cc'])) {
            $message['ccRecipients'] = $this->parseRecipients($data['cc']);
        }

        if (! empty($data['bcc'])) {
            $message['bccRecipients'] = $this->parseRecipients($data['bcc']);
        }

        return $message;
    }

    /**
     * Converts a plain-text compose body (line breaks only, no markup) into safe HTML so
     * Outlook/Graph — which renders message bodies as HTML — preserves the author's line breaks.
     */
    private function htmlFromPlainText(string $text): string
    {
        return nl2br(htmlspecialchars($text, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8'));
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
