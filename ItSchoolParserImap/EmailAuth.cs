using System;
using System.Collections.Generic;
using System.Text;

namespace ItSchoolParserImap
{
    [Serializable]
    public class EmailAuth
    {
        public string Hostname { get; set; }
        public int Port { get; set; }
        /// <summary>
        /// Использование сертификата безопасности
        /// </summary>
        public bool UseSsl { get; set; }
        /// <summary>
        /// Адрес эл почты, например: <c>example@gmail.com</c>
        /// </summary>
        public string Username { get; set; }
        /// <summary>
        /// Пароль от почты
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// Название папки, из которой будут подгружаться письма
        /// </summary>
        public string Folder { get; set; }
        /// <summary>
        /// Текст, с которого начинается ТЕМА письма, должен быть общим для всех искомых писем
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Адрес искомого отправителя 
        /// </summary>
        public string Sender { get; set; }

        public EmailAuth()
        {
            Hostname = "*Введите_адрес_хоста*";
            Port = 143;
            UseSsl = false;
            Username = "*Введите_адрес_электронной_почты*";
            Password = "*Введите_пароль*";
            Folder = "*Введите_название_папки*";
            Title = "*Введите_начало_темы*";
            Sender = "*Введите_адрес_отправителя*";
        }
        public EmailAuth(string hostname, int port, bool useSsl, string username, string password)
        {

        }
    }
}
