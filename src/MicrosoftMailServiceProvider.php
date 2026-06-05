<?php

namespace Nukeflame\MicrosoftMail;

use GuzzleHttp\Client;
use Illuminate\Support\ServiceProvider;

class MicrosoftMailServiceProvider extends ServiceProvider
{
    public function register(): void
    {
        $this->app->singleton(MailRepository::class);

        $this->app->singleton(MicrosoftMailService::class, function ($app) {
            return new MicrosoftMailService(
                $app->make(MailRepository::class),
                new Client(),
            );
        });
    }
}
