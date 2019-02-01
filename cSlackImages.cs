using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace excelParcer
{
    class cSlackImages
    {
        //Место хранения изображения
        public string imagePath { get; set; }
        
        //Имя файла
        public string fName { get; set; }
        
        //Размер файла
        public long fSize { get; set; }

        //Формирование значения на основании XML узла
        public bool FromXMLElement(XmlNode elm)
        {
            //Если элемент на входе 
            if (elm == null)
                return false;
            //Получаем путь к изображению
            this.imagePath = elm.Attributes["image_path"].Value.ToString();
            //Получаем размер файла
            this.fSize = Convert.ToInt32(elm.Attributes["image_size"].Value);
            //Получаем значение элемента
            this.fName = elm.InnerText;

            return true;
        }

        //Преобразование параметров в объект класса XMLElement
        public XmlNode ToXMLNode(XmlDocument slkDoc)
        {
            //Основной элемент
            XmlElement slackImage;
            //Аттрибуты imagePath - путь, imageSize- размер
            XmlAttribute imagePath, imageSize;
            //Текстовые значения для атрибутов
            XmlText imagePathText,          //значение аттрибута path_text
                    imageSizeText,          //значение аттрибута size_text
                    imageName;              //Текстовое значение имени файла
            //Создание корневого элемента
            slackImage = slkDoc.CreateElement("image");
            //Создание аттрибута: путь к картинке
            imagePath = slkDoc.CreateAttribute("image_path");
            //Присваиваем значение для данного аттрибута
            imagePathText = slkDoc.CreateTextNode(this.imagePath);

            //Создание аттрибута
            imageSize = slkDoc.CreateAttribute("image_size");
            //Присваиваем значение "размер" для данного атрибута
            imageSizeText = slkDoc.CreateTextNode(this.fSize.ToString());   
            
            //Задаём значение "имя" для узла
            imageName = slkDoc.CreateTextNode(this.fName);

            //Добавляем текст к аттрибутам
            imagePath.AppendChild(imagePathText);
            imageSize.AppendChild(imageSizeText);

            //Присоединяем атррибуты
            slackImage.Attributes.Append(imagePath);
            slackImage.Attributes.Append(imageSize);
            //Добавляем текст в узел
            slackImage.AppendChild(imageName);
            //Возвращаем элемент
            return slackImage;
        }
        
        //Сравнение двух файлов
        public bool FileCompare(String newFileName)
        {
            //Сравниваем пути.
            //Возвращает TRUE если файлы одинаковые или пути указывают на один и тот же объект            
            if (newFileName == this.imagePath + this.fName)
                return true;
            
            //Инициируем переменные
            //Переменные для записывания и сравнения файлов побайтово
            int fileOldByte, fileNewByte;
            
            //Потоки для открытия файлов
            FileStream fileCurrent, fileNew;
            //Отокрываем текущий и новый файл
            fileCurrent = new FileStream(this.imagePath + this.fName, FileMode.Open);
            fileNew = new FileStream(newFileName, FileMode.Open);

            do
            {
                //Считываем побайтово
                fileOldByte = fileCurrent.ReadByte();
                fileNewByte = fileNew.ReadByte();
            }
            while //Пока байты равны - продолжаем цикл
            ((fileOldByte == fileNewByte) && (fileOldByte != -1));
            //Отключаем потоки
            fileCurrent.Close();
            fileNew.Close();
            //Если при вычитании последние байты дают 0 - файлы полностью идентичны
            return ((fileOldByte - fileNewByte) == 0);
        }
    }
}
