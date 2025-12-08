using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;

namespace WordMarkdownAddIn.Services
{
    public class WordToMarkdownService
    {
        private readonly Application _wordApp;
        private readonly Document _activeDoc;

        public WordToMarkdownService() 
        {
            _wordApp = Globals.ThisAddIn.Application;
            _activeDoc = _wordApp.ActiveDocument;

            if (_activeDoc == null)
                throw new System.Exception("Нет активного документа Word.");
        }
        public List<IWordElement> ExtractDocumentStructure()
        {
            var elements = new List<IWordElement>();

            elements = ExtractParagraphs(elements);
            elements = ExtractTables(elements);


            return elements; 
        
        }

        // 1. Таблицы
        private List<IWordElement> ExtractTables(List<IWordElement> elements) 
        {
            foreach (Table table in _activeDoc.Tables)
            {
                var tableData = new List<List<string>>();

                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    var rowData = new List<string>();
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        string cellText = table.Cell(i, j).Range.Text.TrimEnd('\r', '\a');
                        rowData.Add(cellText);
                    }
                    tableData.Add(rowData);
                }

                elements.Add(new WordTable { Rows = tableData});
            }
            return elements;

        }

        // 2. Параграфы
        private List<IWordElement> ExtractParagraphs(List<IWordElement> elements )
        {
            // Обходим все параграфы
            foreach (Paragraph para in _activeDoc.Paragraphs)
            {
                // Убираем символ конца параграфа
                string text = para.Range.Text.TrimEnd('\r', '\a');
                
                if (string.IsNullOrEmpty(text)) continue;
                //Определяем тип параграфа
                string styleName = para.get_Style().NameLocal;

                if (styleName.Contains("Heading")) // Заголовок
                {
                    elements.Add(new WordParagraph { Text = text, StyleName = styleName });
                }
                else 
                {
                    elements.Add(new WordParagraph { Text = text, StyleName = styleName });
                }
            }
            return elements;

        }

        // 3. Гиперссылки
        private void ExtractHyperlinks()
        {

        }

        //  4. Изображения
        private void ExtractImages()
        {

        }

        // 5. Закладки
        private void ExtractBookmarks()
        {

        }

        // 6. Сноски
        private void ExtractFootnotes()
        {

        }

    }
}
