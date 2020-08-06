using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ImapX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ItSchoolParserImap
{
    class GoogleSpreadsheetHandler
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public List<Message> MessageList { get; }
        public string ApplicationName { get; }
        public string SpreadsheetId { get; }
        public string Sheet { get; }
        public string CredentialFile { get; }
        public SheetsService Service { get; set; }

        public List<Message__c> BufferList { get; set; }
        public List<object> ObjectList { get; set; }
        public List<object> BatchList { get; set; }


        public GoogleSpreadsheetHandler(List<Message> messageList, GoogleAuth googleAuth)
        {
            MessageList = messageList;
            ApplicationName = googleAuth.ApplicationName;
            SpreadsheetId = googleAuth.SpreadsheetId;
            Sheet = googleAuth.Sheet;
            CredentialFile = googleAuth.CredentialFile;
        }

        public int SpreadsheetsConnect()
        {
            GoogleCredential credential;
            try
            {
                using (FileStream fs = new FileStream(CredentialFile, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(fs).CreateScoped(Scopes);
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("Файл .json с учетными данными проекта не найден, проверьте корректность названия");
                return 1;
            }
            catch (IOException)
            {
                Console.WriteLine("Файл .json с учетными данными проекта не найден, проверьте корректность названия");
                return 1;
            }

            Service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer() { 
                HttpClientInitializer = credential, 
                ApplicationName = ApplicationName 
            });

            CreateEntry();
            return 0;
        }

        private void CreateEntry()
        {
            var range = $"{Sheet}!A:M";
            Regex regex = new Regex(@"((<\w{1,2}\/>|<\w{1,2}>|<\/\w>)|\\\w)");
            BufferList = new List<Message__c>();

            foreach (Message msg_i in MessageList)
            {
                BufferList.Add(new Message__c(msg_i.Date, regex.Replace(msg_i.Body.Html, "")));
            }


            /// <summary>
            /// Дата обновления таблицы
            /// </summary>
            if (!(BufferList.Count == 0))
            {
                DateTime currentDate = DateTime.Now;
                var valueRange = new ValueRange();
                BatchList = new List<object>() { currentDate.ToShortDateString(), currentDate.ToShortTimeString() };
                valueRange.Values = new List<IList<object>> { BatchList };
                var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                var appendResponse = appendRequest.Execute();
                BatchList.Clear();
            }

            foreach (var msg_i in BufferList)
            {
                try
                {
                    var valueRange = new ValueRange();
                    BatchList = new List<object>()
                    {
                        $"{msg_i.DateSent.Value.Day}/{msg_i.DateSent.Value.Month}/{msg_i.DateSent.Value.Year}",
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Выбранные курсы:") + 16, msg_i.MessageAsText.Length - msg_i.MessageAsText.IndexOf("Выбранные курсы:") - 16),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Имя: ") + 5, msg_i.MessageAsText.IndexOf("Возраст (для детей)") - msg_i.MessageAsText.IndexOf("Имя: ") - 5 ),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Телефон: ") + 9, msg_i.MessageAsText.IndexOf("Viber: ") - msg_i.MessageAsText.IndexOf("Телефон: ") - 9),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Viber: ") + 7, msg_i.MessageAsText.IndexOf("Telegram: ") - msg_i.MessageAsText.IndexOf("Viber: ") - 7),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Telegram: ") + 10, msg_i.MessageAsText.IndexOf("E-mail: ") - msg_i.MessageAsText.IndexOf("Telegram: ") - 10),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("E-mail: ") + 8, msg_i.MessageAsText.IndexOf("Откуда вы узнали о нас: ") - msg_i.MessageAsText.IndexOf("E-mail: ") - 8),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Откуда вы узнали о нас: ") + 24, msg_i.MessageAsText.IndexOf("Локация: ") - msg_i.MessageAsText.IndexOf("Откуда вы узнали о нас: ") - 24),
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Локация: ") + 9, msg_i.MessageAsText.IndexOf("Дополнительные комментарии / вопросы:") - msg_i.MessageAsText.IndexOf("Локация: ") - 9),
                        "Not specified yet",
                        msg_i.MessageAsText.Substring(msg_i.MessageAsText.IndexOf("Возраст (для детей): ") + 21, msg_i.MessageAsText.IndexOf("Телефон: ") - msg_i.MessageAsText.IndexOf("Возраст (для детей): ") - 21)
                    };

                    /// <summary>
                    /// Эта коллекция будет записываться в строку
                    /// </summary>
                    valueRange.Values = new List<IList<object>> { BatchList };
                    var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    var appendResponse = appendRequest.Execute();
                }
                catch (FormatException)
                {
                    Console.WriteLine("Ошибка в учётных данных Гугл таблиц \"GoogleAuth.xml\"");
                    break;
                }
                catch(Exception ex)
                {
                    Console.WriteLine($"Error in \"GoogleSpreadsheetHandler.cs\": {ex}");
                }
            }
        }
    }
}
