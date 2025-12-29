using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;   // Подключение пространства имён для взаимодействия с Microsoft Word.
using System.Text.RegularExpressions;  // Подключение пространства имён для работы с регулярными выражениями. 

// Объявление пространства имён для сервисов, связанных с преобразованием Word в Markdown.
namespace WordMarkdownAddIn.Services
{
    /// <summary>
    /// Сервис для извлечения текстового содержимого документа Word без форматирования.
    /// Предназначен для случаев, когда в документе Word уже записан текст по правилам Markdown.
    /// </summary>
    public class WordToMarkdownPlainTextService
    {
        private readonly Application _wordApp;   // Приватное поле для хранения ссылки на приложение Word.
        private readonly Document _activeDoc;    // Приватное поле для хранения ссылки на активный документ Word.

        /// <summary>
        /// Инициализирует сервис, получая ссылку на активный документ Word.
        /// </summary>
        /// <exception cref="System.Exception">Выбрасывается, если активный документ отсутствует.</exception>
        public WordToMarkdownPlainTextService()
        {
            _wordApp = Globals.ThisAddIn.Application;  // Получение ссылки на приложение Word через Globals (объект, предоставляемый VSTO).
            _activeDoc = _wordApp.ActiveDocument;       // Приватное поле для хранения ссылки на активный документ Word. 

            // Проверка, существует ли активный документ.
            if (_activeDoc != null) 
            {
                // Если активного документа нет, выбрасывается исключение.
                throw new System.Exception("Нет активного документа Word.");
            }


        }
    
    }        
}

