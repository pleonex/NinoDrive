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
using System.IO;
using System.Reflection;

namespace NinoDrive
{
    public class AuthorizationManager
    {
        private const string ClientId = ".apps.googleusercontent.com";
        private const string Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
        private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";
        private const string DefaultProgramName = "NinoDrive-v1";

        private const string MagicTokenWord = "VELDANA";
        private readonly static System.Text.Encoding Encoding = System.Text.Encoding.ASCII;
        private readonly static string AssemblyPath = 
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private readonly static string TokenFile = Path.Combine(AssemblyPath, "token");

        private static volatile AuthorizationManager instance;
        private static readonly object syncRoot = new object();

        private OAuth2Parameters oauthParams;
        private GOAuth2RequestFactory requestFactory;

        private AuthorizationManager()
        {
            ProgramName = DefaultProgramName;
            GetAuthorization();
        }

        ~AuthorizationManager()
        {
            SaveTokenToFile();
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

        public void InvalidateAuthorization()
        {
            File.Delete(TokenFile);
            GetAuthorization();
        }

        private void GetAuthorization()
        {
            // Get the client ID and secret from environment variables for security.
            var clientId = GetSecretVariable("NINODRIVE_ID");
            var clientSecret = GetSecretVariable("NINODRIVE_SECRET");
            if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret)) {
                Console.WriteLine("ERROR Set NINODRIVE_ID and NINODRIVE_SECRET env vars");
                Environment.Exit(-1);
            }

            // Prepare the OAuth params.
            oauthParams = new OAuth2Parameters();
            oauthParams.ClientId = clientId + ClientId;
            oauthParams.ClientSecret = clientSecret;
            oauthParams.RedirectUri = RedirectUri;
            oauthParams.Scope = Scope;

            // Since it's a factory it's safe to continue update OAuthParams.
            requestFactory = new GOAuth2RequestFactory(null, ProgramName, oauthParams);

            // Try to get the previous token, otherwise request access code.
            if (!ReadTokenFromFile())
                AskAccessCode();
        }

        private static string GetSecretVariable(string varName)
        {
            // Maybe we could also store them in a json file like Google allows.
            return Environment.GetEnvironmentVariable(varName);
        }

        private bool ReadTokenFromFile()
        {
            if (!File.Exists(TokenFile))
                return false;

            byte[] data = File.ReadAllBytes(TokenFile);
            byte[] key  = Encoding.GetBytes(GetSecretVariable("NINODRIVE_SECRET"));

            for (int i = 0; i < data.Length; i++)
                data[i] ^= key[i % key.Length];

            string[] content = Encoding.GetString(data).Split('\n');

            bool isValid = (content.Length == 3) && (content[0] == MagicTokenWord);
            if (isValid) {
                oauthParams.AccessToken  = content[1];
                oauthParams.RefreshToken = content[2];
            }

            return isValid;
        }

        private void SaveTokenToFile()
        {
            if (oauthParams == null || string.IsNullOrEmpty(oauthParams.AccessToken) ||
                string.IsNullOrEmpty(oauthParams.RefreshToken))
                return;

            string content = MagicTokenWord + '\n' + oauthParams.AccessToken + '\n' +
                             oauthParams.RefreshToken;
            byte[] key  = Encoding.GetBytes(GetSecretVariable("NINODRIVE_SECRET"));
            byte[] data = Encoding.GetBytes(content);

            for (int i = 0; i < data.Length; i++)
                data[i] ^= key[i % key.Length];

            File.WriteAllBytes(TokenFile, data);
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

