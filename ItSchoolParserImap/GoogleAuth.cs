using System;
using System.Collections.Generic;
using System.Text;

namespace ItSchoolParserImap
{
    [Serializable]
    /// <summary>
    /// Класс хранит данные, касающиеся конкретной таблицы
    /// </summary>
    public class GoogleAuth
    {
        /// <summary>
        /// A name of google application
        /// </summary>
        public string ApplicationName { get; set; }
        /// <summary>
        /// An id of spreadsheet carried in URL
        /// </summary>
        public string SpreadsheetId { get; set; }
        /// <summary>
        /// A name of the spreadsheets list we're working with
        /// </summary>
        public string Sheet { get; set; }
        /// <summary>
        /// A name of credentials file we're fetching when aplication created.
        /// Credential for authorizing calls using OAuth 2.0. It is a convenience wrapper
        /// that allows handling of different types of credentials
        /// </summary>
        public string CredentialFile { get; set; }

        public GoogleAuth()
        {
            ApplicationName = "*Введите_название_приложения*";
            SpreadsheetId = "*Введите_айди_таблицы*";
            Sheet = "*Введите_название_листа*";
            CredentialFile = "*Введите_название_файла_учётных_данных*";
        }
    }
}
