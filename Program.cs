using System;
using System.Threading.Tasks;

using CsvHelper;
using System.IO;
using System.Text.RegularExpressions;
using System.Configuration;

namespace OutlookDataLabellingTool
{
    class Program
    {
        private OutlookDataExtractor _dataExtractor = null;
        static void Main(string[] args)
        {
            var clientId = ConfigurationManager.AppSettings.Get("clientId");
            var labellingQuestion = ConfigurationManager.AppSettings.Get("labellingQuestion");
            var csvFilename = ConfigurationManager.AppSettings.Get("csvFilename");
            var startDate = DateTime.Parse(ConfigurationManager.AppSettings.Get("startDate"));
            var endDate = DateTime.Parse(ConfigurationManager.AppSettings.Get("endDate"));

            var program = new Program(clientId);
            program.LabelSentItems(labellingQuestion, startDate, endDate, csvFilename).Wait();

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }

        internal Program(string clientId)
        {
            _dataExtractor = new OutlookDataExtractor(clientId);
        }

        private async Task LabelSentItems(string labellingQuestion, DateTime startDate, DateTime endDate, string csvFilePath)
        {
            string[] scopes = { "Mail.Read" };

            // Convert to ISO 8601 date/time string format.  Note: the "o" format string does not work properly with C# 6.0 interpolated expressions (ie: $"{o:dateTimeVar}")
            var startDateString = startDate.ToUniversalTime().ToString("o");
            var endDateString = endDate.ToUniversalTime().ToString("o");
            // ToDo: add recipients
            var url = $"https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages?$select=id,subject,uniqueBody,sentDateTime&$filter=(sentDateTime ge {startDateString}) and (sentDateTime le {endDateString})";

            var accessToken = await _dataExtractor.GetAccessToken(scopes);

            Console.WriteLine($"For each email message, please answer the question as Y/N:  {labellingQuestion}\n");

            await WriteDataToCsvFile(csvFilePath, async (CsvWriter csvWriter) =>
            {
                await _dataExtractor.GetOutlookDataAsText(accessToken, url, (dynamic message) =>
                {
                    // skip auto-generated calendar event messages (ie: responses to invites)
                    if (message["@odata.type"] == "#microsoft.graph.eventMessage")
                        return Task.CompletedTask;

                    string messageContent = message.uniqueBody.content;

                    // skip forwarded messages
                    if (messageContent.StartsWith("Your meeting was forwarded"))
                        return Task.CompletedTask;

                    // skip auto-generated file attachment messages
                    if (messageContent.StartsWith("Your message is ready to be sent with the following file or link attachments"))
                        return Task.CompletedTask;

                    // Remove http links inserted during Outlook's HTML to text conversion
                    messageContent = Regex.Replace(messageContent, @"<(http|https)[^\s]+>", string.Empty);

                    // remove default mobile footers
                    var androidMessage = "Get Outlook for Android\r\n";
                    if (messageContent.EndsWith(androidMessage))
                        messageContent = messageContent.Remove(messageContent.Length - androidMessage.Length);

                    var iosMessage = "Get Outlook for iOS\r\n";
                    if (messageContent.EndsWith(iosMessage))
                        messageContent = messageContent.Remove(messageContent.Length - iosMessage.Length);

                    // skip empty messages
                    if (messageContent.Length == 0)
                        return Task.CompletedTask;

                    var travelIntent = false;
                    Console.WriteLine("\n===>>> Start message.");
                    Console.WriteLine(messageContent);
                    Console.Write($"\n===>>> End message. {labellingQuestion} (Y/N)");
                    var key = Console.ReadKey();
                    Console.WriteLine();
                    if (key.Key == ConsoleKey.Y)
                        travelIntent = true;

                    string id = message.id;
                    string sentDate = message.sentDateTime;
                    csvWriter.WriteField(sentDate);
                    csvWriter.WriteField(id);
                    csvWriter.WriteField(messageContent);
                    csvWriter.WriteField<bool>(travelIntent);
                    csvWriter.NextRecord();
                    return Task.CompletedTask;
                });
            });
        }


        //public async Task LabelCalendarLocations(DateTime startDate, DateTime endDate, string csvFilePath)
        //{
        //    string[] scopes = { "Mail.Read" };

        //    // Convert to ISO 8601 date/time string format.  Note: the "o" format string does not work properly with C# 6.0 interpolated expressions (ie: $"{o:dateTimeVar}")
        //    var startDateString = startDate.ToUniversalTime().ToString("o");
        //    var endDateString = endDate.ToUniversalTime().ToString("o");

        //    var url = $"https://graph.microsoft.com/v1.0/me/events?$select=subject,body,bodyPreview,organizer,attendees,start,end,location&$filter=(sentDateTime ge {startDateString}) and (sentDateTime le {endDateString})";

        //    var accessToken = await _dataExtractor.GetAccessToken(scopes);

        //    Console.WriteLine("For each calendar event location, please answer the question:  Doe this location contain an address?");

        //    await WriteDataToCsvFile(csvFilePath, async (CsvWriter csvWriter) =>
        //    {
        //        await _dataExtractor.GetOutlookDataAsText(accessToken, url,  (dynamic calendarEvent) =>
        //        {
        //            string id = calendarEvent.id;
        //            string location = calendarEvent.location;

        //            if (calendarEvent.location.Length > 0)
        //            {
        //                csvWriter.WriteField(id);
        //                Console.WriteLine(id);

        //                csvWriter.WriteField(location);
        //                Console.WriteLine(location);
        //                csvWriter.NextRecord();
        //            }
        //            return Task.CompletedTask;
        //        });
        //    });
        //}

        private async Task WriteDataToCsvFile(string path, Func<CsvWriter, Task> handler)
        {
            using (var streamWriter = new StreamWriter(path))
            {
                // CsvWriter writes data to CSV files per RFC 4180
                using (var csv = new CsvWriter(streamWriter))
                {
                    await handler(csv);
                }
            }
        }

    }
}
