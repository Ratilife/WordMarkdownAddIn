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
            if (_activeDoc == null) 
            {
                // Если активного документа нет, выбрасывается исключение.
                throw new System.Exception("Нет активного документа Word.");
            }
        }

        /// <summary>
        /// Извлекает текстовое содержимое документа Word без форматирования.
        /// </summary>
        /// <returns>Текстовое содержимое документа с нормализованными переносами строк.</returns>
        public string ExtractPlainText() 
        {
            try
            {
                //Получаем текстовое содержимое документа
                string text = _activeDoc.Content.Text;

                //Нормализуем переносы строк
                text = NormalizeLineBreaks(text);

                // Нормализуем множественные пустые строки
                text = NormalizeEmptyLines(text);

                // Удаление пробельных символов и переносов строк из начала строки.
                text = text.TrimStart('\n', '\r', ' ', '\t');
                // Удаление пробельных символов и переносов строк из конца строки.
                text = text.TrimEnd('\n', '\r', ' ', '\t');
                return text;
            }
            catch (Exception ex) 
            {
                // Логируем ошибку для отладки
                System.Diagnostics.Debug.WriteLine($"Ошибка при извлечении текста: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Нормализует переносы строк в тексте.
        /// Заменяет все варианты переносов строк (Windows, Mac, Unix) на единый формат (\n).
        /// </summary>
        /// <param name="text">Исходный текст.</param>
        /// <returns>Текст с нормализованными переносами строк.</returns>
        private string NormalizeLineBreaks(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;    // Возврат исходного текста, если он null или пуст.

            // Заменяем Windows-переносы (\r\n) на \n
            text = text.Replace("\r\n", "\n");

            // Заменяем Mac-переносы (\r) на \n
            text = text.Replace("\r", "\n");

            // Удаляем символы конца параграфа Word (\a - звонок)
            text = text.Replace("\a", "");

            return text;
        }

        /// <summary>
        /// Нормализует множественные пустые строки подряд.
        /// Заменяет 3 и более переносов строк подряд на максимум 2.
        /// </summary>
        /// <param name="text">Исходный текст.</param>
        /// <returns>Текст с нормализованными пустыми строками.</returns>
        private string NormalizeEmptyLines(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // Заменяем 3 и более переносов строк подряд на 2 переноса
            // Используем цикл while для обработки всех случаев (например, \n\n\n\n\n)
            while (text.Contains("\n\n\n"))
            {
                text = text.Replace("\n\n\n", "\n\n");
            }

            return text;
        }


    }
}

