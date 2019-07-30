using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace testOutlookAddIn
{
    static class FileProcess
    {
        private static string PathToMainFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        /// <summary>
        /// Сохранение данных в файл .txt по пути "Мои документы"
        /// </summary>
        /// <param name="nameOfFile">Имя файла</param>
        /// <param name="body">Содержимое файла</param>
        public static void SaveToFile(string nameOfFile, string body)
        {
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(File.Open(PathToMainFile + $@"\{nameOfFile}.txt", FileMode.Create)))
                {
                    streamWriter.Write(body);
                }
            }
            catch
            {
                MessageBox.Show("Ошибка записи", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    }
}
