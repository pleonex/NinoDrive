// Copyright (C) 2015 zkarts
// Copyright (C) 2016 pleonex
using System.Windows.Forms;
using System;
using Google.GData.Client;

namespace NinoDrive
{
    public static class Program
    {
        private const string ClientId = ".apps.googleusercontent.com";
        private const string Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
        private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

        [STAThreadAttribute] // Needed to write the link to the clipboard
        public static void Main()
        {
            // Get the authorization
            var authorization = GetAuthorization();

            // Get the services and start working.
            GoogleToXML go = new GoogleToXML("NinoDrive-v1", authorization);
            go.PrepareService();
            while (true)
                go.FromSpreadsheet();
        }

        private static OAuth2Parameters GetAuthorization()
        {
            // Get the client ID and secret from environment variables for security.
            var clientId = Environment.GetEnvironmentVariable("NINODRIVE_ID") + ClientId;
            var clientSecret = Environment.GetEnvironmentVariable("NINODRIVE_SECRET");

            // Prepare the OAuth params.
            OAuth2Parameters oauthParams = new OAuth2Parameters();
            oauthParams.ClientId = clientId;
            oauthParams.ClientSecret = clientSecret;
            oauthParams.RedirectUri = RedirectUri;
            oauthParams.Scope = Scope;

            // Ask to get the authorization code. Give the URL and ask for the code.
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(oauthParams);
            Clipboard.SetText(authorizationUrl);
            Console.WriteLine("The following link has been copied to your clipboard.");
            Console.WriteLine("Please paste that link in a browser and authorize");
            Console.WriteLine("yourself on the webpage. Once that is complete, type");
            Console.WriteLine("your access code and press Enter to continue...");
            Console.WriteLine(authorizationUrl);

            Console.Write("Access code: ");
            oauthParams.AccessCode = Console.ReadLine();
            Console.WriteLine("-----------------------------------------");
            Console.WriteLine();

            return oauthParams;
        }
    }
}