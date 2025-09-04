using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordMarkdownAddIn.Controls
{
    public class TaskPaneControl: UserControl
    {
        private readonly WebView2 _webView;
        private readonly Services.MarkdownRenderService _renderer;
        private string _latestMarkdown = string.Empty;  // Инициализация пустой строкой  Локальный кэш для быстрого доступа
        private bool _coreReady = false;

        public TaskPaneControl() 
        {
            _renderer = new Services.MarkdownRenderService();
            _webView = new WebView2
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_webView);                     // Добавляет WebView2 на UserControl   Controls - это коллекция всех дочерних элементов управления
            Load += OnLoadAsync;                        // Подписываем метод OnLoadAsync на событие Load
        }

        private async void OnLoadAsync(object sender, EventArgs e) 
        {
        
        }



    }
}
