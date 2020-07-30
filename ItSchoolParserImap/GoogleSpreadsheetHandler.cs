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

        public List<string> BufferList { get; set; }
        public List<object> ObjectList { get; set; }
        public List<string> BatchList { get; set; }


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
            var range = $"{Sheet}!A:J"; //var range = $"{Sheet}!A:I";
            Regex regex = new Regex(@"((<\w{1,2}\/>|<\w{1,2}>|<\/\w>)|\\\w)");
            BufferList = new List<string>();

            /// <summary>
            /// Коллекция типа Message приводится к типу string. По причине инвариантности
            /// обобщений (они не принимают участия в иерархии наследования типов) наиболее
            /// простой способ привести типы - переписать каждую переменную отдельно.
            /// </summary>
            foreach (var msg_i in MessageList)
            {
                BufferList.Add(msg_i.Body.Html);
            }

            /// <summary>
            /// Из сообщения убираются лишние тэги, знаки табуляции, переносы строки и пр.
            /// </summary>
            for (int i = 0; i < BufferList.Count; i++)
            {
                BufferList[i] = regex.Replace(BufferList[i], "");
            }



            /// <summary>
            /// Дата обновления таблицы
            /// </summary>
            if (!(BufferList.Count == 0))
            {
                DateTime currentDate = DateTime.Now;
                var valueRange = new ValueRange();
                ObjectList = new List<object>() { currentDate.ToShortDateString(), currentDate.ToShortTimeString() };
                valueRange.Values = new List<IList<object>> { ObjectList };
                var appendRequest = Service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                var appendResponse = appendRequest.Execute();
                ObjectList.Clear();


            }

            foreach (var msg_i in BufferList)
            {
                try
                {
                    var valueRange = new ValueRange();
                    ObjectList = new List<object>();
                    BatchList = new List<string>()
                    {
                        msg_i.Substring(msg_i.IndexOf("Имя: ") + 5, msg_i.IndexOf("Возраст (для детей)") - msg_i.IndexOf("Имя: ") - 5 ),
                        msg_i.Substring(msg_i.IndexOf("Возраст (для детей): ") + 21, msg_i.IndexOf("Телефон: ") - msg_i.IndexOf("Возраст (для детей): ") - 21/**/),
                        msg_i.Substring(msg_i.IndexOf("Телефон: ") + 9, msg_i.IndexOf("Viber: ") - msg_i.IndexOf("Телефон: ") - 9),
                        msg_i.Substring(msg_i.IndexOf("Viber: ") + 7, msg_i.IndexOf("Telegram: ") - msg_i.IndexOf("Viber: ") - 7),
                        msg_i.Substring(msg_i.IndexOf("Telegram: ") + 10, msg_i.IndexOf("E-mail: ") - msg_i.IndexOf("Telegram: ") - 10),
                        msg_i.Substring(msg_i.IndexOf("E-mail: ") + 8, msg_i.IndexOf("Откуда вы узнали о нас: ") - msg_i.IndexOf("E-mail: ") - 8),
                        msg_i.Substring(msg_i.IndexOf("Откуда вы узнали о нас: ") + 24, msg_i.IndexOf("Локация: ") - msg_i.IndexOf("Откуда вы узнали о нас: ") - 24),
                        msg_i.Substring(msg_i.IndexOf("Локация: ") + 9, msg_i.IndexOf("Дополнительные комментарии / вопросы:") - msg_i.IndexOf("Локация: ") - 9),
                        msg_i.Substring(msg_i.IndexOf("Дополнительные комментарии / вопросы:") + 38, msg_i.IndexOf("Выбранные курсы:") - msg_i.IndexOf("Дополнительные комментарии / вопросы:") - 38),
                        msg_i.Substring(msg_i.IndexOf("Выбранные курсы:") + 16, msg_i.Length - msg_i.IndexOf("Выбранные курсы:") - 16)
                    };

                    foreach (var msg_j in BatchList)
                    {
                        ObjectList.Add(msg_j);
                    }

                    /// <summary>
                    /// Эта коллекция будет записываться в строку
                    /// </summary>
                    valueRange.Values = new List<IList<object>> { ObjectList };
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
