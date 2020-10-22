using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Message = Google.Apis.Gmail.v1.Data.Message;

namespace GmailTask
{
    class Program
    {
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "Gmail API .NET Quickstart";

        static void Main(string[] args)
        {

            UserCredential credential;
            using (var stream =
                new FileStream(@"C:\Users\Akash Thakker\Documents\Gmail Task\GmailTask\credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None
                    ).Result;
                //Console.WriteLine("Credential file saved to: " + credPath);
            }
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            var result = MessageBox.Show("Start Deleting Mails?", "Gmail Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                string text = "";
                var searches = System.IO.File.ReadAllText(@"C:\Users\Akash Thakker\Documents\Gmail Task\GmailTask\Gmail.txt").Split(',');
                foreach (var search in searches)
                {
                    List<Message> messages = ListMessages(service, "me", search);
                    text += search + " = " + messages.Count() + Environment.NewLine;
                    for (int i = 0; i < messages.Count(); i++)
                    {
                        if (messages.Count() != 0)
                        {
                            DeleteMessage(service, "me", messages[i].Id.ToString());
                            if (i == messages.Count() - 1)
                            {

                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                MessageBox.Show(text);
                List<string> Moresearchs = new List<string>();
                var searchtext = Interaction.InputBox("DO you want to delete some others? if yes then enter the name with ',' (or) else click enter to complete", "Gmail", "").Split(',');
                if (searchtext != null && searchtext.Count() != 0 && searchtext[0] != "")
                {
                    foreach (var searchitem in searchtext)
                    {
                        List<Message> messages = ListMessages(service, "me", searchitem);
                        for (int i = 0; i < messages.Count(); i++)
                        {
                            DeleteMessage(service, "me", messages[i].Id.ToString());
                            if (i > messages.Count() - 1)
                            {
                            }
                        }
                    }
                }
                MessageBox.Show("Delete Process Successful");
            }
            var resultmails = MessageBox.Show("Move selected mails?", "Mail Moving", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultmails == DialogResult.Yes)
            {
                ModifyMessages(service);
            }
            MessageBox.Show("Service Ends");
        }
        public static void DeleteMessage(GmailService service, String userId, String messageId)
        {
            try
            {
                service.Users.Messages.Trash(userId, messageId).Execute();
            }
            catch (Exception e)
            {
                //Console.WriteLine("An error occurred: " + e.Message);
            }
        }

        public static List<Message> ListMessages(GmailService service, String userId, String query)
        {
            List<Message> result = new List<Message>();
            UsersResource.MessagesResource.ListRequest request = service.Users.Messages.List(userId);
            request.Q = query;
            do
            {
                try
                {
                    ListMessagesResponse response = request.Execute();
                    if (response.Messages != null)
                        result.AddRange(response.Messages);
                    request.PageToken = response.NextPageToken;
                }
                catch (Exception e)
                {
                    // Console.WriteLine("An error occurred: " + e.Message);
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return result;
        }
        public static Message ModifyMessages(GmailService service)
        {
            string text = "";
            List<String> labelsToAdd = new List<String>();
            List<String> labelsToRemove = new List<String>();
            var Mails = System.IO.File.ReadAllText(@"C:\Users\Akash Thakker\Documents\Gmail Task\GmailTask\GmailTransfer.txt").Split(',');
            ListLabelsResponse ListLabels = service.Users.Labels.List("me").Execute();
            foreach (var Mailsitem in Mails)
            {
                if (Mailsitem != "" || Mailsitem != null)
                {
                    int count = 0;
                    var mailsearch = Mailsitem.Split('-');
                    List<Message> listmessages = ListMessages(service, "me", mailsearch[0]);
                    foreach (var messagesvalue in listmessages)
                    {
                        count = 0;
                        labelsToAdd.Clear();
                        labelsToRemove.Clear();
                        var addlabelsvalue = ListLabels.Labels.Where(x => x.Name == mailsearch[1]).First().Id;
                        var labelids = service.Users.Messages.Get("me", messagesvalue.Id).Execute();
                        if (!labelids.LabelIds.Contains(addlabelsvalue))
                        {
                            for (int i = 0; i < labelids.LabelIds.Count(); i++)
                            {
                                labelsToRemove.Add(labelids.LabelIds[i].ToString());
                            }
                            labelsToAdd.Add(addlabelsvalue);

                            count++;
                            ModifyMessageRequest mods = new ModifyMessageRequest();
                            mods.AddLabelIds = labelsToAdd;
                            mods.RemoveLabelIds = labelsToRemove;
                            service.Users.Messages.Modify(mods, "me", messagesvalue.Id).Execute();
                        }

                    }
                    text += mailsearch[0] + " moved = " + count + Environment.NewLine;
                }

            }
            MessageBox.Show(text);
            return null;
        }
    }
}
