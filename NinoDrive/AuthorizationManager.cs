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
            requestFactory = new GOAuth2RequestFactory(null, ProgramName, oauthParams);
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

            // Try to get the previous token
            if (!ReadTokenFromFile())
                AskAccessCode();
            else if (!TestAuthorization())
                RefreshAuthorization();
        }

        private bool ReadTokenFromFile()
        {
            throw new NotImplementedException();
        }

        private bool TestAuthorization()
        {
            throw new NotImplementedException();
        }

        private void RefreshAuthorization()
        {
            OAuthUtil.RefreshAccessToken(oauthParams);
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

