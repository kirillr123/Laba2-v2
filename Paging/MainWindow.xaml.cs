using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections;
using System.ComponentModel;
using Paging;
using System.Net;
using System.IO;
using Microsoft.Win32;

namespace GridPagingExample
{
    public partial class MainWindow : Window
    {
        private PagingCollectionView _cview;
        public NoteDB db;

        public MainWindow()
        {
            InitializeComponent();
            var xlsxPath = Directory.GetCurrentDirectory() + "\\temp.xlsx";//Laba2\Laba2\bin\Release\temp.xlsx
            var dataPath = Directory.GetCurrentDirectory() + "\\data.xml";//Laba2\Laba2\bin\Release\data.xml
            if(File.Exists(dataPath))
            {
                db = NoteDB.DeSerializeObject<NoteDB>(dataPath);
            }
            else
            {
                DownloadFile(xlsxPath);
                db = new NoteDB();
                db.Parse(xlsxPath);
            }

            NoteDB.SerializeObject<NoteDB>(db, dataPath);

            this._cview = new PagingCollectionView(db.list
                ,
                20
            );
            this.DataContext = this._cview;

        }

        //смена страниц
        private void OnNextClicked(object sender, RoutedEventArgs e)
        {
            this._cview.MoveToNextPage();
        }

        private void OnPreviousClicked(object sender, RoutedEventArgs e)
        {
            this._cview.MoveToPreviousPage();
        }
        private void OnSaveClicked(object sender, RoutedEventArgs e)
        {
            // Displays a SaveFileDialog so the user can save the Image  
            // assigned to Button2.  
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dlg.FileName;
                File.Copy(Directory.GetCurrentDirectory() + "\\temp.xlsx", filename);

            }



        }
        //Обновление БД
        private void OnRenewClicked(object sender, RoutedEventArgs e)
        {
            //скачиваем файл и сверяем каждую строку в xlsx с каждой записью в БД
            var xlsxPath = Directory.GetCurrentDirectory() + "\\temp.xlsx";//Laba2\Laba2\bin\Release\temp.xlsx
            var dataPath = Directory.GetCurrentDirectory() + "\\data.xml";//Laba2\Laba2\bin\Release\data.xml
            DownloadFile(xlsxPath);
            db.AnotherParse(xlsxPath);
            NoteDB.SerializeObject<NoteDB>(db, dataPath);


            this._cview = new PagingCollectionView(db.list
                ,
                20
            );
            this.DataContext = this._cview;

        }
        
        //реализация текстбокса
        private void OnShowClicked(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(userInput.Text, out int id) )
            {
                MessageBox.Show("Введите пожалуйста номер записи! Он должен быть положительным целым числом!");
            }
            else if(id<1)
            {
                MessageBox.Show("Идентификатор должен быть положительным!");
            }
            else
            {
                try
                {
                    Note note = db.list.Find(x => x.Id == "УБИ."+id);
                    if (note!=null)
                    {
                        MessageBox.Show("Нужная запись была найдена "+ Environment.NewLine+"Идентификатор: "
                            +note.Id+ Environment.NewLine+"Наименование УБИ: "+ note.Name+ Environment.NewLine + "Описание: " + note.Description + Environment.NewLine
                            + "Источник угрозы: " + note.Source + Environment.NewLine + "Объект воздействия: " + note.Object + Environment.NewLine+ "Нарушение конфиденциальности: "+
                            note.IsConfidential + Environment.NewLine + "Нарушение целостности: " + note.IsInteger + Environment.NewLine + "Нарушение доступности: "
                            + note.IsAccesible + Environment.NewLine);
                    }
                    else
                    {
                        MessageBox.Show("Запись не найдена!!");
                    }
                    
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Произошла ошибка! Возможно вы указали неправильный идентификатор! Ошибка: " + ex.Message);
                }
            }
        }
        private void TextBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            TextBox txtBox = sender as TextBox;
            if (txtBox.Text == "ID записи")
                txtBox.Text = string.Empty;
        }
        private void DownloadFile(string xlsxPath)
        {
            MessageBox.Show("Сейчас произойдет загрузка данных с сервера.");
            using (WebClient client = new WebClient())
            {
                Uri uri = new Uri("https://bdu.fstec.ru/documents/files/thrlist.xlsx");
                client.DownloadFile(uri, xlsxPath);
            }
        }
    }
}