using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows;
using ClosedXML.Excel;
using System.Xml;
using System.Xml.Serialization;

namespace Paging
{  
    [Serializable]
    public class NoteDB
    {
        public List<Note> list = new List<Note>();
        public void Parse(string file)
        {
            try
            {
                //open the excel using openxml sdk  

                var workbook = new XLWorkbook(file);
                var rows = workbook.Worksheet(1).RangeUsed().RowsUsed().Skip(2);

                foreach (var row in rows)
                {
                    Note note = new Note
                    {
                        Id = "УБИ."+row.Cell(1).Value.ToString(),
                        Name = row.Cell(2).Value.ToString(),
                        Description = row.Cell(3).Value.ToString(),
                        Source = row.Cell(4).Value.ToString(),
                        Object = row.Cell(5).Value.ToString(),
                        IsConfidential = Convert.ToBoolean(row.Cell(6).Value),
                        IsInteger = Convert.ToBoolean(row.Cell(7).Value),
                        IsAccesible = Convert.ToBoolean(row.Cell(8).Value)
                    };
                    list.Add(note);
                }
            }
            catch (Exception Ex)
            {

                MessageBox.Show(Ex.Message);
            }
        }
        public void AnotherParse(string file)
        {
            NoteDB newdb = new NoteDB();
            List<string> strings = new List<string>();
            string s="";
            newdb.Parse(file);
            int i = 0;
            int count = 0;
            try
            {
                foreach (var item in newdb.list)
                {
                    if (i < list.Count)
                    {
                        if (!item.Equals(list[i]))
                        {
                            count++;
                            s = s + "Запись номер: " + item.Id + Environment.NewLine;
                            if (!item.Name.Equals(list[i].Name))
                            {
                                s=s+"Наименование. БЫЛО: " +list[i].Name + Environment.NewLine+ "СТАЛО: "+ item.Name+Environment.NewLine;
                            }
                            if (!item.Description.Equals(list[i].Description))
                            {
                                s = s + "Описание. БЫЛО: " + list[i].Description + Environment.NewLine + "СТАЛО: " + item.Description + Environment.NewLine;
                            }
                            if (!item.Source.Equals(list[i].Source))
                            {
                                s = s + "Источник. БЫЛО: " + list[i].Source + Environment.NewLine + "СТАЛО: " + item.Source + Environment.NewLine;
                            }
                            if (!item.Object.Equals(list[i].Object))
                            {
                                s = s + "Объект. БЫЛО: " + list[i].Object + Environment.NewLine + "СТАЛО: " + item.Object + Environment.NewLine;
                            }
                            if (item.IsConfidential != list[i].IsConfidential)
                            {
                                s = s + "Нарушение конфиденциальности. БЫЛО: " + list[i].IsConfidential + Environment.NewLine + "СТАЛО: " + item.IsConfidential + Environment.NewLine;
                            }
                            if (item.IsInteger != list[i].IsInteger)
                            {
                                s = s + "Нарушение целостности. БЫЛО: " + list[i].IsInteger + Environment.NewLine + "СТАЛО: " + item.IsInteger + Environment.NewLine;
                            }
                            if (item.IsAccesible != list[i].IsAccesible)
                            {
                                s = s + "Нарушение доступа. БЫЛО: " + list[i].IsAccesible + Environment.NewLine + "СТАЛО: " + item.IsAccesible + Environment.NewLine;
                            }
                        }
                        i++;
                    }
                    else
                    {
                        count++;
                    }
                }
                MessageBox.Show("Обновление произошло успешно! Количество изменений:" + (count+list.Count-i).ToString());
                if (s != "")
                {
                    MessageBox.Show("Детальный обзор изменений:" + Environment.NewLine + s);
                }
                else if((count+list.Count-i)!=0)
                {
                    MessageBox.Show("Изменений существующих записей не произошло! Были добавлены новые записи или удалены старые.");
                }
                this.list = newdb.list;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка в обновлении:" + ex.Message);
                    
            }
        }

        /// <summary>
        /// Serializes an object.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serializableObject"></param>
        /// <param name="fileName"></param>
        public static void SerializeObject<T>(T serializableObject, string fileName)
        {
            if (serializableObject == null) { return; }

            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
                using (MemoryStream stream = new MemoryStream())
                {
                    serializer.Serialize(stream, serializableObject);
                    stream.Position = 0;
                    xmlDocument.Load(stream);
                    xmlDocument.Save(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);//Log exception here
            }
        }


        /// <summary>
        /// Deserializes an xml file into an object list
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static T DeSerializeObject<T>(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) { return default(T); }

            T objectOut = default(T);

            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(fileName);
                string xmlString = xmlDocument.OuterXml;

                using (StringReader read = new StringReader(xmlString))
                {
                    Type outType = typeof(T);

                    XmlSerializer serializer = new XmlSerializer(outType);
                    using (XmlReader reader = new XmlTextReader(read))
                    {
                        objectOut = (T)serializer.Deserialize(reader);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);//Log exception here
            }

            return objectOut;
        }
    }
}
