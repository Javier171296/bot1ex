// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
namespace SampleAADV2Bot.Dialogs
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using AuthBot;
    using AuthBot.Dialogs;
    using AuthBot.Models;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Connector;
    using System.Net.Http;
    using Newtonsoft.Json.Linq;

    [Serializable]
    public class ActionDialog : IDialog<string>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task TokenSample(IDialogContext context,String mess)
        {
            //endpoint v2
            var accessToken = await context.GetAccessToken(AuthSettings.Scopes);

            if (string.IsNullOrEmpty(accessToken))
            {
                return;
            }

            //search logic
            //await context.PostAsync($"Your access token is: {accessToken}");

            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);

            // get page content as html https://graph.microsoft.com/beta/me/onenote/pages/1-4a1562d50ce64ecf8cc62d0d19f97d79!8-ff7b7482-b307-4166-abf6-d9c0462b85ab/content

            // get all pages https://graph.microsoft.com/v1.0/me/onenote/pages

            // get user´s basic information https://graph.microsoft.com/v1.0/me/

            using (var response = await client.GetAsync("https://graph.microsoft.com/v1.0/me/onenote/pages?$select=title,links"))
            {
                if (response.IsSuccessStatusCode)
                {

                    JObject json = JObject.Parse(await response.Content.ReadAsStringAsync());


                    //var a = json.SelectToken("value").Value<String>();

                    var a = json.ToString();

                    await context.PostAsync($"This is the list of all pages that you have in your OneNote: "+a);
                }
                else
                {
                    await context.PostAsync($"Sorry but I did not find any page in your OneNote " + response.IsSuccessStatusCode + "   " + response.ReasonPhrase.ToString() );
                }
            }


            context.Wait(MessageReceivedAsync);
        }

        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            if (message.Text == "logon")
            {
                //endpoint v2
                if (string.IsNullOrEmpty(await context.GetAccessToken(AuthSettings.Scopes)))
                {
                    await context.Forward(new AzureAuthDialog(AuthSettings.Scopes), this.ResumeAfterAuth, message, CancellationToken.None);
                }
                else
                {
                    context.Wait(MessageReceivedAsync);
                }
            }
            else if (message.Text == "hi")
            {
                await context.PostAsync("Hello, I´m happy to work with you");

                context.Wait(this.MessageReceivedAsync);
            }
            else if (message.Text == "give me information")
            {
                await TokenSample(context,message.Text);
            }
            else if (message.Text == "logout")
            {
                await context.Logout();
                context.Wait(this.MessageReceivedAsync);
            }
            else
            {
                context.Wait(MessageReceivedAsync);
            }
        }

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            await context.PostAsync(message);
            context.Wait(MessageReceivedAsync);
        }
    }
}


//*********************************************************
//
//AuthBot, https://github.com/microsoftdx/AuthBot
//
//Copyright (c) Microsoft Corporation
//All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:




// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.




// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
//*********************************************************
