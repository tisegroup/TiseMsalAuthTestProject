import { ApplicationConfig, provideBrowserGlobalErrorListeners } from '@angular/core';
import { provideRouter } from '@angular/router';
import { routes } from './app.routes';
import { MsalInterceptorConfiguration, MsalGuardConfiguration, MSAL_GUARD_CONFIG, MSAL_INSTANCE, MSAL_INTERCEPTOR_CONFIG, MsalBroadcastService, MsalGuard, MsalInterceptor, MsalService } from '@azure/msal-angular';
import { IPublicClientApplication, PublicClientApplication, BrowserCacheLocation, InteractionType } from '@azure/msal-browser';
import { provideHttpClient, withInterceptorsFromDi, HTTP_INTERCEPTORS } from '@angular/common/http';

// 1. MSAL Instance Configuration
export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: 'f44037cf-a5dc-421a-a698-d2a96889a780',
      authority: 'https://login.microsoftonline.com/29fbf026-3349-4188-96eb-18a387320c2a',
      redirectUri: 'http://localhost:4200',
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
    },
  });
}

// 2. Interceptor Configuration (Maps your .NET API URL to its Entra ID Scope)
export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();
  
  // Use a wildcard '*' so all sub-routes under /api/ receive the token
  protectedResourceMap.set('https://localhost:44323/api/*', [
    'api://51c79005-ab59-478c-93bc-fef9f55680d2/user_impersonation',
  ]);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap,
  };
}

// 3. Route Guard Configuration
export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: {
      scopes: ['api://YOUR_DOTNET_API_CLIENT_ID/user_impersonation'],
    },
  };
}

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideRouter(routes),
    provideHttpClient(withInterceptorsFromDi()),
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true,
    },
    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory,
    },
    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory,
    },
    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory,
    },
    MsalService,
    MsalGuard,
    MsalBroadcastService,
  ]
};
