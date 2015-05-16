using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace winform_Calendar
{
    public partial class Form1 : Form
    {
        private IList<string> scopes = new List<string>();
        private CalendarService service;
        public Form1()
        {
            InitializeComponent();

            try
            {
                scopes.Add(CalendarService.Scope.Calendar);

                UserCredential credential;

                //oauth2
                using (FileStream stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                       GoogleClientSecrets.Load(stream).Secrets, scopes, "user", CancellationToken.None,
                       new FileDataStore("")).Result;
                }

                // Create Calendar Service.
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "app",
                });

                var list = service.CalendarList.List().Execute().Items;
                var list2 = service.Calendars.Get("messboy000@gmail.com").Execute();

                // Define parameters of request.
                //Calendar ID 就是去網頁裡面找設定
                EventsResource.ListRequest request = service.Events.List("messboy000@gmail.com");
                request.TimeMin = DateTime.Now;
                request.ShowDeleted = false;
                request.SingleEvents = true;
                request.MaxResults = 30;
                request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                textBox1.Text += string.Format("Upcoming events: \r\n");
                var events = request.Execute();
                if (events.Items.Count > 0)
                {
                    foreach (var eventItem in events.Items)
                    {
                        string when = eventItem.Start.DateTime.ToString();
                        if (String.IsNullOrEmpty(when))
                        {
                            when = eventItem.Start.Date;
                        }
                        textBox1.Text += string.Format("{0} ({1}) \r\n", eventItem.Summary, when);
                    }
                }
                else
                {
                    textBox1.Text += string.Format("No upcoming events found.");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }

        



        }
    }
}
