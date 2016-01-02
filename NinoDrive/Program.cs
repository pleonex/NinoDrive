// Copyright (C) 2015 zkarts
// Copyright (C) 2016 pleonex
using System.Windows.Forms;
using System;
using Google.GData.Client;

namespace NinoDrive
{
    public static class Program
    {
        private const string ClientId = "YOUR_CLIENT_ID.apps.googleusercontent.com";
        private const string ClientSecret = "ADD YOUR OWN CLIENT_SECRET";
        private const string Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
        private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

        [STAThreadAttribute] // Needed to write the link to the clipboard
        public static void Main(string[] args)
        {
            // Prepare the OAuth params.
            OAuth2Parameters oauthParams = new OAuth2Parameters();
            oauthParams.ClientId = ClientId;
            oauthParams.ClientSecret = ClientSecret;
            oauthParams.RedirectUri = RedirectUri;
            oauthParams.Scope = Scope;

            // Ask to get the authorization code. Give the URL and ask for the code.
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(oauthParams);
            Clipboard.SetText(authorizationUrl);
            Console.WriteLine("A link has been copied to your clipboard. Please paste that"
            + " link in a browser and authorize yourself on the webpage.  Once that is"
            + " complete, type your access code and press enter to continue...");
            oauthParams.AccessCode = Console.ReadLine();

            // Get the services and start working.
            GoogleToXML go = new GoogleToXML("NinoDrive-v1", oauthParams);

            Console.WriteLine("-----------------------------------------");
            go.PrepareService();
            while (true)
                go.FromSpreadsheet();
        }
    }
}