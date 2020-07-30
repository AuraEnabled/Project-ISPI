using ImapX;
using ImapX.Collections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
//using System.Runtime.Serialization;
using System.Xml.Serialization;
using System.Text;
using ImapX.Enums;

namespace ItSchoolParserImap
{
    class EmailHandler
    {
        public string Hostname { get; set; }
        public int Port { get; set; }
        public bool UseSsl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string Folder { get; set; }
        public string Title { get; set; }
        public string Sender { get; set; }
        public DateTime StdDate { get; set; }
        public List<Message> allNewMessages { get; set; }

        private DateTime? latestDate;
        private string latestMessageText;
        private Message__c latestMessage;

        public EmailHandler(EmailAuth emailAuth)
        {
            Hostname = emailAuth.Hostname;
            Port = emailAuth.Port;
            UseSsl = emailAuth.UseSsl;
            Username = emailAuth.Username;
            Password = emailAuth.Password;
            Folder = emailAuth.Folder;
            Title = emailAuth.Title;
            Sender = emailAuth.Sender;
            StdDate = new DateTime(1971
                , 1, 1, 15, 0, 0);
        }

        public List<Message> ObtainMessages()
        {
            using (ImapClient client = new ImapClient(Hostname, Port, UseSsl, false))
            {
                if (!client.Connect())
                {
                    Console.WriteLine("Не подключился к серверу");
                    return allNewMessages;
                }

                if (!client.Login(Username, Password))
                {
                    Console.WriteLine("Неверный логин или пароль");
                    return allNewMessages;
                }
                else
                {
                    if (client.IsConnected) Console.WriteLine("Connection established");


                    /// <summary>
                    /// Сообщения скачиваются из нужной папки
                    /// </summary>

                    string query = "ALL";
                    MessageFetchMode mode = (MessageFetchMode)(-1);
                    client.Folders["INBOX"].SubFolders[Folder].Messages.Download(query, mode, 50);
                    int messageCount = client.Folders.Inbox.SubFolders[Folder].Messages.Count();
                    MessageCollection messages = client.Folders.Inbox.SubFolders[Folder].Messages;
                    XmlSerializer formatter = new XmlSerializer(typeof(Message__c));
                    allNewMessages = new List<Message>();
                    messageCount--;

                    /// <summary>
                    /// Начальное присваивание
                    /// </summary>
                    for (int i = messageCount; i >= 0; i-- )
                    {
                        if (messages[i].From.Address.Equals(Sender)
                            &&messages[i].Subject.StartsWith(Title)
                            )
                        {
                            latestDate = messages[i].Date;
                            latestMessageText = messages[i].Body.Html;
                            break;
                        }
                        else if (i == 0)
                        {
                            Console.WriteLine("Новых сообщений нет");
                            return allNewMessages;
                        }
                    }

                    /// <summary>
                    /// Поиск наиболее позднего
                    /// </summary>
                    for (int i = messageCount; i >= 1; i--)
                    {
                        if (latestDate < messages[i - 1].Date
                            &&messages[i - 1].From.Address.Equals(Sender)
                            &&messages[i - 1].Subject.StartsWith(Title)
                            )
                        {
                            latestDate = messages[i - 1].Date;
                            latestMessageText = messages[i - 1].Body.Html;
                        }
                    }

                    /// <summary>
                    /// Десериализация сообщения из прошлой сессии
                    /// </summary>
                    try
                    {
                        using(FileStream fs = new FileStream("message.xml", FileMode.Open))
                        {
                            Message__c newMessage = (Message__c)formatter.Deserialize(fs);
                            if (newMessage.DateSent < latestDate)
                            {
                                latestMessage = new Message__c(latestDate, latestMessageText);
                            }
                            else
                            {
                                return allNewMessages;
                            }

                            for (int i = messageCount; i >= 0; i--)
                            {
                                if(messages[i].Date > newMessage.DateSent 
                                    && messages[i].From.Address.Equals(Sender) 
                                    && messages[i].Subject.StartsWith(Title)
                                    )
                                {
                                    allNewMessages.Add(messages[i]);
                                }
                            }
                        }
                    }
                    catch (FileNotFoundException)
                    {
                        Console.WriteLine("Анализ всех входящих писем");
                        latestMessage = new Message__c(latestDate, latestMessageText);

                        for (int i = messageCount; i >= 0; i--)
                        {
                            if (messages[i].Date > StdDate
                                &&messages[i].From.Address.Equals(Sender)
                                &&messages[i].Subject.StartsWith(Title)
                                )
                            {
                                allNewMessages.Add(messages[i]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error in \"EmailHandler\": {ex}");
                    }

                    /// <summary>
                    /// Сериализация последнего сообщения
                    /// </summary>
                    using (FileStream fs = new FileStream("message.xml", FileMode.Create))
                    {
                        formatter.Serialize(fs, latestMessage);
                    }


                    return allNewMessages;
                }
            } 
        }
    }
}
