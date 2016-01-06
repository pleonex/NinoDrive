//
//  AuthorizationManager.cs
//
//  Authors:
//       pleonex <benito356@gmail.com>
//       zkarts
//
//  Copyright (c) 2016 pleonex
//
//  This program is free software: you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation, either version 3 of the License, or
//  (at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with this program.  If not, see <http://www.gnu.org/licenses/>.
using System;
using Google.GData.Client;
using System.Net;

namespace NinoDrive
{
    public class AuthorizationManager
    {
        private const string ClientId = ".apps.googleusercontent.com";
        private const string Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
        private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";
        private const string DefaultProgramName = "NinoDrive-v1";
            
        private static volatile AuthorizationManager instance;
        private static readonly object syncRoot = new object();

        private OAuth2Parameters oauthParams;
        private GOAuth2RequestFactory requestFactory;

        private AuthorizationManager()
        {
            ProgramName = DefaultProgramName;

            GetAuthorization();
        }

        public static AuthorizationManager Instance {
            get {
                if (instance == null) {
                    lock (syncRoot) {
                        if (instance == null)
                            instance = new AuthorizationManager();
                    }
                }

                return instance;
            }
        }

        public string ProgramName {
            get;
            private set;
        }

        public void UpdateService(Service service)
        {
            service.RequestFactory = requestFactory;
        }

        private void GetAuthorization()
        {
            // Get the client ID and secret from environment variables for security.
            var clientId = Environment.GetEnvironmentVariable("NINODRIVE_ID") + ClientId;
            var clientSecret = Environment.GetEnvironmentVariable("NINODRIVE_SECRET");
 
            // Prepare the OAuth params.
            oauthParams = new OAuth2Parameters();
            oauthParams.ClientId = clientId;
            oauthParams.ClientSecret = clientSecret;
            oauthParams.RedirectUri = RedirectUri;
            oauthParams.Scope = Scope;

            // Since it's a factory it's safe to continue update OAuthParams.
            requestFactory = new GOAuth2RequestFactory(null, ProgramName, oauthParams);

            // Try to get the previous token, otherwise request access code.
            if (!ReadTokenFromFile())
                AskAccessCode();
        }

        private bool ReadTokenFromFile()
        {
            // We only need to store the latest token and the refresh token.
            throw new NotImplementedException();
        }

        private bool TestAuthorization()
        {
            var service = new Google.GData.Spreadsheets.SpreadsheetsService(ProgramName);
            UpdateService(service);

            bool valid = true;
            try {
                service.Query(new Google.GData.Spreadsheets.SpreadsheetQuery());
            } catch (GDataRequestException) {
                valid = false;
            }

            return valid;
        }

        private void RefreshAuthorization()
        {
            try {
                OAuthUtil.RefreshAccessToken(oauthParams);
            } catch (WebException) {
                AskAccessCode();
            }
        }

        private void AskAccessCode()
        {
            // Ask to get the authorization code. Give the URL and ask for the code.
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(oauthParams);
            Console.WriteLine("Please open the following link in a browser and authorize");
            Console.WriteLine("yourself on the webpage. Once that is complete, type");
            Console.WriteLine("your access code and press Enter to continue.");
            Console.WriteLine(authorizationUrl);

            Console.Write("Access code: ");
            oauthParams.AccessCode = Console.ReadLine();
            Console.WriteLine("-----------------------------------------");
            Console.WriteLine();

            OAuthUtil.GetAccessToken(oauthParams);
        }
    }
}

