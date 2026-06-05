<?php

namespace Nukeflame\MicrosoftMail;

use App\Models\UserMailToken;

class MailRepository
{
    public function findByUser(int $userId): ?UserMailToken
    {
        return UserMailToken::query()->where('user_id', $userId)->first();
    }

    public function store(int $userId, array $data): UserMailToken
    {
        return UserMailToken::query()->updateOrCreate(
            ['user_id' => $userId],
            [
                'access_token'  => $data['access_token'],
                'refresh_token' => $data['refresh_token'],
                'expires_at'    => $data['expires_at'],
                'scope'         => $data['scope'] ?? null,
                'token_type'    => $data['token_type'] ?? 'Bearer',
            ]
        );
    }

    public function update(UserMailToken $token, array $data): void
    {
        $token->forceFill([
            'access_token'  => $data['access_token'],
            'refresh_token' => $data['refresh_token'] ?? $token->refresh_token,
            'expires_at'    => $data['expires_at'],
        ])->save();
    }

    public function delete(int $userId): void
    {
        UserMailToken::query()->where('user_id', $userId)->delete();
    }

    public function isConnected(int $userId): bool
    {
        return UserMailToken::query()->where('user_id', $userId)->exists();
    }
}
