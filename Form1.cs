using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Xml;


//Подключаем excel
using Excel = Microsoft.Office.Interop.Excel;

namespace excelParcer
{
    public partial class Form1 : Form
    {
        List<cSlack> slacks;
        public Form1()
        {
            InitializeComponent();
            slacks = new List<cSlack>();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Открываем диалог.
            //Если отказываем - выход
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            //Присвоение имени файла
            this.label7.Text = openFileDialog1.FileName;

        }
        
        //Поиск следующей позиции для разбора
        private List<int> getNextLineNum(string pattern, Excel.Worksheet src)
        {
            //Список для возврата
            List<int> res = new List<int> { };
            Excel.Range currentRange = null,
                        firstFind = null;


            object misValue = System.Reflection.Missing.Value;
            int i = 1;
            Excel.Range colRange = src.Cells[1, 6000];
            string itemAddress;
            string s;
            Regex regex = new Regex(@"\d{1}\w{1}\d{4}\s{4}");
            MatchCollection matches;// = regex.Matches(s);
            
            while ( i < 6000)
            {
                itemAddress = "A" + i.ToString();
                if(src.Range[itemAddress].Value == null)
                {
                    i++;
                    continue;
                }
                s = src.Range[itemAddress].Value.ToString();
                matches = regex.Matches(s);
                if (matches.Count > 0)
                {
                    res.Add(i);
                }
                

                i++;
            }


            
            /*currentRange = colRange.Find(pattern, misValue, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                    Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, true, misValue, misValue);
            res.Add(currentRange.Row);
            Excel.Range foundCell   = currentRange;
            Excel.Range firstResult = foundCell;*/
/*
            while (foundCell != null)
            {
                
                Excel.Range foundTemp = foundCell;
                //foundCell = currentRange.FindNext(foundTemp);
                foundCell = colRange.FindNext(foundTemp);
                res.Add(foundCell.Row);

                if (foundCell.Address == firstResult.Address)
                    foundCell = null;
            }*/
            /*            
                        Excel.Range resRange = colRange.Find(pattern);
                        Excel.Range fountTemp;
                        res.Add(resRange.Row);
                        int currentRow = resRange.Row;
                        while (currentRow < 6000)
                        {
                            fountTemp = resRange;
                            resRange = colRange.FindNext(fountTemp);
                            currentRow = resRange.Row;
                            res.Add(currentRow);
                        }*/
            //resRange.Row
            //Выбрать столбец А

            //src.Range["A1:A1000"].
            //Сравнить значение текущей строки с шаблоном посредством регулярного выражения
            //Если выражение совпадение найдено - вернуть номер строки
            //Если совпадение не найдено - крутим цикл дальше.
            //При достижении нижней границы цикла, если ничего не найдено - значит в предыдущий раз был найден последний элемент
            return res;
        }
        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Открываем диалог.
            //Если отказываем - выход
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            //Присвоение имени файла
            string fileName = openFileDialog1.FileName;

            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();

            excelApp.Visible = true;
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(fileName);
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];

            /*cSlack slk = new cSlack();
            string val = wrkSheet.Range["A4"].Value;
            slk.id = val.ToString();

            richTextBox1.AppendText(slk.id);*/

            string[] input = { "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L" };
            List<string> colls = new List<string>(input);
            string cellNumberText, cellNumberValue;

            XmlDocument xDoc = new XmlDocument();
            
            xDoc.Load("slacks.xml");
            XmlElement xRoot = xDoc.DocumentElement;

            

            for(int j = 4; j <= wrkSheet.Rows.CurrentRegion.EntireRow.Count; j++)
            //for (int j = 4; j <= 10; j++)
            { 
                // создаем новый элемент slack
                XmlElement userElem = xDoc.CreateElement("slack");
                // создаем атрибут id
                XmlAttribute slack_id = xDoc.CreateAttribute("id");
                // создаем атрибут to_fit
                XmlAttribute to_fit = xDoc.CreateAttribute("to_fit");
                // создаем атрибут axle_side
                XmlAttribute axle_side = xDoc.CreateAttribute("axle_side");
                // создаем атрибут other_side
                XmlAttribute other_side = xDoc.CreateAttribute("other_side");
                // создаем атрибут vehicle_type
                XmlAttribute vehicle_type = xDoc.CreateAttribute("vehicle_type");
                // создаем атрибут type
                XmlAttribute slack_type = xDoc.CreateAttribute("slack_type");
                // создаем атрибут trand_mark
                XmlAttribute trand_mark = xDoc.CreateAttribute("trand_mark");


                // создаем элементы crossess
                XmlElement crossess = xDoc.CreateElement("crossess");
                // создаём элемент hole_params
                XmlElement hole_params = xDoc.CreateElement("hole_params");
                // создаём элемент slack_params
                XmlElement slack_params = xDoc.CreateElement("slack_params");

                // создаем текстовые значения для элементов и атрибута
                // артикул
                XmlText slack_id_text   =   xDoc.CreateTextNode(wrkSheet.Range["A"+j].Value);

                CSlackType type_mark = new CSlackType(wrkSheet.Range["A" + j].Value);
                XmlText slack_type_text = xDoc.CreateTextNode(type_mark.SlactType);
                XmlText trand_mark_text = xDoc.CreateTextNode(type_mark.TrandMark);
                // марка транспортного средства/производитель оси
                XmlText to_fit_text     =   xDoc.CreateTextNode(wrkSheet.Range["M"+j].Value == null ? "" : wrkSheet.Range["M"+j].Value);
                // сторона
                XmlText axle_side_text = xDoc.CreateTextNode("");
                // другая сторона
                XmlText other_side_text = xDoc.CreateTextNode("");
                // тип транспортного средства
                XmlText vehicle_type_text = xDoc.CreateTextNode(wrkSheet.Range["N"+j].Value == null ? "" : wrkSheet.Range["N"+j].Value);

                //Аттрибуты
                slack_id.AppendChild(slack_id_text);
                to_fit.AppendChild(to_fit_text);
                axle_side.AppendChild(axle_side_text);
                other_side.AppendChild(other_side_text);
                vehicle_type.AppendChild(vehicle_type_text);
                slack_type.AppendChild(slack_type_text);
                trand_mark.AppendChild(trand_mark_text);

                //Кроссы - crossess
                for (int i = 0; i < colls.Count(); i++)
                {
                    cellNumberText = colls[i] + "1";
                    cellNumberValue = colls[i] + j;
                    XmlElement cross = xDoc.CreateElement("cross");
                    XmlAttribute cross_descr = xDoc.CreateAttribute("descr");
                    //TODO:: оптимизировать выборку пустых значений для ускорения формирования общего списка
                    XmlText cross_text = xDoc.CreateTextNode(wrkSheet.Range[cellNumberValue].Value == null ? "" : wrkSheet.Range[cellNumberValue].Value.ToString());
                    XmlText cross_descr_text = xDoc.CreateTextNode(wrkSheet.Range[cellNumberText].Value == null ? "" : wrkSheet.Range[cellNumberText].Value);

                    cross_descr.AppendChild(cross_descr_text);
                    cross.Attributes.Append(cross_descr);
                    cross.AppendChild(cross_text);
                    crossess.AppendChild(cross);
                    cross_text = null;
                    cross_descr_text = null;
                    cross_descr = null;
                    cross = null;

                }

                userElem.Attributes.Append(slack_id);
                userElem.Attributes.Append(to_fit);
                userElem.Attributes.Append(axle_side);
                userElem.Attributes.Append(other_side);
                userElem.Attributes.Append(vehicle_type);
                userElem.Attributes.Append(slack_type);
                userElem.Attributes.Append(trand_mark);

                //crossess.
                userElem.AppendChild(crossess);
                userElem.AppendChild(hole_params);
                userElem.AppendChild(slack_params);
                xRoot.AppendChild(userElem);
                
                richTextBox1.AppendText(slack_id.Value+" | "+ type_mark.SlactType + " | " + type_mark.TrandMark + "\r\n");
                //Зачистка переменных

                slack_id = null;
                to_fit = null;
                axle_side = null;
                other_side = null;
                vehicle_type = null;
                crossess = null;
                hole_params = null;
                slack_params = null;
                slack_type = null;
                trand_mark = null;

                slack_id_text = to_fit_text = axle_side_text = other_side_text = vehicle_type_text = slack_type_text = trand_mark_text = null;
                userElem = null;
                

            }
            //добавляем узлы
            xDoc.Save("slk.xml");
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
        }

        private void xmlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Подгружаем xml
            /*if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                //return;

            string fName = openFileDialog1.FileName;
            XmlDocument doc = new XmlDocument();
            doc.Load(fName);

            XmlElement xRoot = doc.DocumentElement;

           // xRoot.
           foreach(XmlNode XmlSlack in xRoot.ChildNodes) { 
                cSlack slk = new cSlack();
                if (slk.LoadFrom(XmlSlack))
                {
                    richTextBox1.AppendText(slk.id.ToString() + " | " + slk.slackType + " | " + slk.trendMark + "\r\n");
                    slacks.Add(slk);
                }
                slk = null;
            }*/
            /*try
            {
                // SECTION 1. Create a DOM Document and load the XML data into it.
                XmlDocument dom = new XmlDocument();
                dom.Load(fName);

                // SECTION 2. Initialize the TreeView control.
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add(new TreeNode(dom.DocumentElement.Name));
                TreeNode tNode = new TreeNode();
                tNode = treeView1.Nodes[0];

                // SECTION 3. Populate the TreeView with the DOM nodes.
                AddNode(dom.DocumentElement, tNode);
                treeView1.ExpandAll();
            }
            catch (XmlException xmlEx)
            {
                MessageBox.Show(xmlEx.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }*/
        }

     

        private void Form1_Load(object sender, EventArgs e)
        {
            XmlDocument doc = new XmlDocument();
            //doc.Load("slk.xml");
            doc.Load("slk2.xml");

            XmlElement xRoot = doc.DocumentElement;

            // xRoot.
            foreach (XmlNode XmlSlack in xRoot.ChildNodes)
            {
                cSlack slk = new cSlack();
                if (slk.LoadFrom(XmlSlack))
                {
                    richTextBox1.AppendText(slk.id.ToString() + " | " + slk.slackType + " | " + slk.trendMark + "\r\n");
                    slacks.Add(slk);
                }
                slk = null;
            }
        }

        private void excelCatalogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();

            excelApp.Visible = true;
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(this.label7.Text);
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];
            
            /**
             * Запускаем поиск по подгруженному файлу
             */
            int i = 0, tmp090 = 0, rowIndex = 0;
            String val;
            foreach(cSlack slk in this.slacks)
            {

                Excel.Range currentRange, cRg;

                cRg = wrkSheet.Columns["A:A"];
                currentRange = cRg.Find(slk.id);//, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext
                //

                if (currentRange != null)
                {
                    richTextBox1.AppendText("Артикул " + slk.id + " найден в ячейке " + currentRange.Row.ToString()+". Загрузка данных в объект\r\n");
                    //Комплект поставки
                    rowIndex = currentRange.Row + 3;
                    val = wrkSheet.Range["K"+ rowIndex.ToString()].Value;
                    slk.suppliedWith = val.Substring(14, val.Length - 14).TrimStart(' ');
                    //Сторона
                    rowIndex = currentRange.Row + 4;
                    val = wrkSheet.Range["K" + rowIndex.ToString()].Value;
                    slk.axleSide = val.Substring(9, val.Length - 9).TrimStart(':');
                    //Другая сторона
                    val = wrkSheet.Range["T" + rowIndex.ToString()].Value;
                    slk.otherSide = val.Substring(11, val.Length - 11).TrimStart(' ');

                    rowIndex = currentRange.Row + 6;
                    //Смещение
                    val = (wrkSheet.Range["K" + rowIndex.ToString()].Value != null ? wrkSheet.Range["K" + rowIndex.ToString()].Value.ToString() : "0");
                    Int32.TryParse(val.Trim(), out tmp090);
                    slk.offset = tmp090;


                    //Наклон
                    //            double inc = wrkSheet.Range["N7"].Value;
                    val = (wrkSheet.Range["N" + rowIndex.ToString()].Value != null ? wrkSheet.Range["N" + rowIndex.ToString()].Value.ToString() : "0");
                    Int32.TryParse(val.Trim(), out tmp090);
                    slk.offset = tmp090;

                    //Угол поводка
                    val = wrkSheet.Range["R" + rowIndex.ToString()].Value != null ? wrkSheet.Range["R" + rowIndex.ToString()].Value.ToString() : "";
                    slk.controlArmAngle = val.Trim();
                    //Int32.TryParse(val.Trim(), out cSlack.controlArmAngle);

                    //Тип поводка: AP, QF
                    val = wrkSheet.Range["V" + rowIndex.ToString()].Value != null ? wrkSheet.Range["V" + rowIndex.ToString()].Value.ToString() : "";
                    slk.controlArmType = val.Trim(); ;

                    //Количество зубьев
                    val = wrkSheet.Range["Y" + rowIndex.ToString()].Value != null ? wrkSheet.Range["Y" + rowIndex.ToString()].Value.ToString() : "0";

                    Int32.TryParse(val.Trim(), out tmp090);

                    slk.splineTeeth = tmp090;

                   //Вилочная втулка: 14.2,
                   val = wrkSheet.Range["AC" + rowIndex.ToString()].Value.ToString();
                    slk.clevisBush = val.Trim();

                    //Расстояние между отверстиями
                    int tmp1 = 0, tmp2 = 0, tmp3 = 0, tmp4 = 0, tmp5 = 0, tmp6 = 0, tmp7 = 0;
                    rowIndex = currentRange.Row + 8;
                    Int32.TryParse((wrkSheet.Range["K" + rowIndex.ToString()].Value != null ? wrkSheet.Range["K" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp1);
                    Int32.TryParse((wrkSheet.Range["N" + rowIndex.ToString()].Value != null ? wrkSheet.Range["N" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp2);
                    Int32.TryParse((wrkSheet.Range["Q" + rowIndex.ToString()].Value != null ? wrkSheet.Range["Q" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp3);
                    Int32.TryParse((wrkSheet.Range["T" + rowIndex.ToString()].Value != null ? wrkSheet.Range["T" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp4);
                    Int32.TryParse((wrkSheet.Range["W" + rowIndex.ToString()].Value != null ? wrkSheet.Range["W" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp5);

                    //string str1 = wrkSheet.Range["Z9"].Value;
                    Int32.TryParse((wrkSheet.Range["Z" + rowIndex.ToString()].Value != null ? wrkSheet.Range["Z" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp6);

                    //str1 = wrkSheet.Range["AC9"].Value;
                    Int32.TryParse((wrkSheet.Range["AC" + rowIndex.ToString()].Value != null ? wrkSheet.Range["AC" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp7);
                    //val = wrkSheet.Range["T5"].Value;
                    slk.holeCentreSize = new List<int> { tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, tmp7 };
                    /*slk.holeCentreSize.Add(tmp2);
                    slk.holeCentreSize.Add(tmp3);
                    slk.holeCentreSize.Add(tmp4);
                    slk.holeCentreSize.Add(tmp5);
                    slk.holeCentreSize.Add(tmp6);
                    slk.holeCentreSize.Add(tmp7);*/

                }
                else
                {
                    richTextBox2.AppendText("Ошибка: Артикул " + slk.id + " НЕ НАЙДЕН!\r\n");
                }
            }

            /*
            List<int> rowsNumbers = this.getNextLineNum("                                                   ", wrkSheet);
            richTextBox1.AppendText(rowsNumbers.Count.ToString() + "\r\n");
            //MessageBox.Show(rowsNumbers.Count.ToString());
            //MessageBox.Show(wrkSheet.Range["A1"].Value);
            slacks = new List<cSlack>() { };

            string val = wrkSheet.Range["A1"].Value;

            cSlack cSlack = new cSlack();
            //ID
            cSlack.id = val.Substring(0, 6);
            //Применяемость
            cSlack.toFit = val.Substring(6, val.Length - 6).TrimStart(' ');
            //Haldex
            val = wrkSheet.Range["K2"].Value;
            cSlack.haldexId.Add(val.Substring(7, val.Length - 7).TrimStart(' '));
            //OEM
            val = wrkSheet.Range["K3"].Value;
            cSlack.oemId.Add(val.Substring(4, val.Length - 4).TrimStart(' '));
            //Комплект поставки
            val = wrkSheet.Range["K4"].Value;
            cSlack.suppliedWith = val.Substring(14, val.Length - 14).TrimStart(' ');
            //Сторона
            val = wrkSheet.Range["K5"].Value;
            cSlack.axleSide = val.Substring(9, val.Length - 9).TrimStart(':');
            //Другая сторона
            val = wrkSheet.Range["T5"].Value;
            cSlack.otherSide = val.Substring(11, val.Length - 11).TrimStart(' ');

            //Смещение
            val = wrkSheet.Range["K7"].Value.ToString();
            //Int32.TryParse(val.Trim(), cSlack.offset);


            //Наклон
            //            double inc = wrkSheet.Range["N7"].Value;
            val = wrkSheet.Range["N7"].Value.ToString();
            //Int32.TryParse(val.Trim(), cSlack.inclination);

            //Угол поводка
            val = wrkSheet.Range["R7"].Value;
            cSlack.controlArmAngle = val.Trim();
            //Int32.TryParse(val.Trim(), out cSlack.controlArmAngle);

            //Тип поводка: AP, QF
            val = wrkSheet.Range["V7"].Value;
            cSlack.controlArmType = val.Trim(); ;

            //Количество зубьев
            val = wrkSheet.Range["Y7"].Value.ToString();
            //Int32.TryParse(val.Trim(), out cSlack.splineTeeth);

            //Вилочная втулка: 14.2,
            val = wrkSheet.Range["AC7"].Value.ToString();
            cSlack.clevisBush = val.Trim();

            //Расстояние между отверстиями
            int tmp1 = 0, tmp2 = 0, tmp3 = 0, tmp4 = 0, tmp5 = 0, tmp6 = 0, tmp7 = 0;

            Int32.TryParse(wrkSheet.Range["K9"].Value.ToString().Trim(), out tmp1);
            Int32.TryParse(wrkSheet.Range["N9"].Value.ToString().Trim(), out tmp2);
            Int32.TryParse(wrkSheet.Range["Q9"].Value.ToString().Trim(), out tmp3);
            Int32.TryParse(wrkSheet.Range["T9"].Value.ToString().Trim(), out tmp4);
            Int32.TryParse(wrkSheet.Range["W9"].Value.ToString().Trim(), out tmp5);

            //string str1 = wrkSheet.Range["Z9"].Value;
            Int32.TryParse((wrkSheet.Range["Z9"].Value != null ? wrkSheet.Range["Z9"].Value.ToString().Trim() : "0"), out tmp6);

            //str1 = wrkSheet.Range["AC9"].Value;
            Int32.TryParse((wrkSheet.Range["AC9"].Value != null ? wrkSheet.Range["AC9"].Value.ToString().Trim() : "0"), out tmp7);
            val = wrkSheet.Range["T5"].Value;
            cSlack.holeCentreSize = new List<int> { tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, tmp7 };
            //cSlack.holeCentreSize.AddRange();
            cSlack.holeCentreSize.Add(tmp2);
            cSlack.holeCentreSize.Add(tmp3);
            cSlack.holeCentreSize.Add(tmp4);
            cSlack.holeCentreSize.Add(tmp5);
            cSlack.holeCentreSize.Add(tmp6);
            cSlack.holeCentreSize.Add(tmp7);*/

            //Изображение
            //val = wrkSheet.Range["T5"].Value;
            /*cSlack.imagesNames = new List<string>();
            //cSlack.imagesNames.Add(cSlack.);


            richTextBox1.AppendText(cSlack.id + "\r\n");
            richTextBox1.AppendText(cSlack.toFit + "\r\n");
            richTextBox1.AppendText(cSlack.haldexId[0].ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.oemId[0].ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.suppliedWith.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.axleSide.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.otherSide.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.offset.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.inclination.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.controlArmAngle.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.controlArmType + "\r\n");
            richTextBox1.AppendText(cSlack.splineTeeth.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.clevisBush.ToString() + "\r\n");
            richTextBox1.AppendText(cSlack.holeCentreSize[0].ToString() + ", " +
                                    cSlack.holeCentreSize[1].ToString() + ", " +
                                    cSlack.holeCentreSize[2].ToString() + ", " +
                                    cSlack.holeCentreSize[3].ToString() + ", " +
                                    cSlack.holeCentreSize[4].ToString() + ", " +
                                    cSlack.holeCentreSize[5].ToString() + ", " +
                                    cSlack.holeCentreSize[6].ToString() + ", " + "\r\n");*/
        }

        private void excelApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Формирование списка рычагов на основании файла EAC Certification List June 2018
            //Открываем диалог.
            //Если отказываем - выход
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            //Присвоение имени файла
            string fileName = openFileDialog1.FileName;

            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();

            excelApp.Visible = true;
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(fileName);
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];

            /*cSlack slk = new cSlack();
            string val = wrkSheet.Range["A4"].Value;
            slk.id = val.ToString();

            richTextBox1.AppendText(slk.id);*/

            string[] input = { "B", "C"};
            List<string> colls = new List<string>(input);
            string cellNumberText, cellNumberValue;

            XmlDocument xDoc = new XmlDocument();

            xDoc.Load("slacks.xml");
            XmlElement xRoot = xDoc.DocumentElement;



            for (int j = 2; j <= wrkSheet.Rows.CurrentRegion.EntireRow.Count; j++)
            //for (int j = 4; j <= 10; j++)
            {
                // создаем новый элемент slack
                XmlElement userElem = xDoc.CreateElement("slack");
                // создаем атрибут id
                XmlAttribute slack_id = xDoc.CreateAttribute("id");
                // создаем атрибут to_fit
                XmlAttribute to_fit = xDoc.CreateAttribute("to_fit");
                // создаем атрибут axle_side
                XmlAttribute axle_side = xDoc.CreateAttribute("axle_side");
                // создаем атрибут other_side
                XmlAttribute other_side = xDoc.CreateAttribute("other_side");
                // создаем атрибут vehicle_type
                XmlAttribute vehicle_type = xDoc.CreateAttribute("vehicle_type");
                // создаем атрибут type
                XmlAttribute slack_type = xDoc.CreateAttribute("slack_type");
                // создаем атрибут trand_mark
                XmlAttribute trand_mark = xDoc.CreateAttribute("trand_mark");


                // создаем элементы crossess
                XmlElement crossess = xDoc.CreateElement("crossess");
                // создаём элемент hole_params
                XmlElement hole_params = xDoc.CreateElement("hole_params");
                // создаём элемент slack_params
                XmlElement slack_params = xDoc.CreateElement("slack_params");

                // создаем текстовые значения для элементов и атрибута
                // артикул
                XmlText slack_id_text = xDoc.CreateTextNode(wrkSheet.Range["A" + j].Value);

                CSlackType type_mark = new CSlackType(wrkSheet.Range["A" + j].Value);
                XmlText slack_type_text = xDoc.CreateTextNode(type_mark.SlactType);
                XmlText trand_mark_text = xDoc.CreateTextNode(type_mark.TrandMark);
                // марка транспортного средства/производитель оси
                XmlText to_fit_text = xDoc.CreateTextNode(wrkSheet.Range["D" + j].Value == null ? "" : wrkSheet.Range["D" + j].Value);
                // сторона
                XmlText axle_side_text = xDoc.CreateTextNode("");
                // другая сторона
                XmlText other_side_text = xDoc.CreateTextNode("");
                // тип транспортного средства
                XmlText vehicle_type_text = xDoc.CreateTextNode("");

                //Аттрибуты
                slack_id.AppendChild(slack_id_text);
                to_fit.AppendChild(to_fit_text);
                axle_side.AppendChild(axle_side_text);
                other_side.AppendChild(other_side_text);
                vehicle_type.AppendChild(vehicle_type_text);
                slack_type.AppendChild(slack_type_text);
                trand_mark.AppendChild(trand_mark_text);

                //Кроссы - crossess
                for (int i = 0; i < colls.Count(); i++)
                {
                    cellNumberText = colls[i] + "1";
                    cellNumberValue = colls[i] + j;
                    XmlElement cross = xDoc.CreateElement("cross");
                    XmlAttribute cross_descr = xDoc.CreateAttribute("descr");
                    //TODO:: оптимизировать выборку пустых значений для ускорения формирования общего списка
                    XmlText cross_text = xDoc.CreateTextNode(wrkSheet.Range[cellNumberValue].Value == null ? "" : wrkSheet.Range[cellNumberValue].Value.ToString());
                    XmlText cross_descr_text = xDoc.CreateTextNode(wrkSheet.Range[cellNumberText].Value == null ? "" : wrkSheet.Range[cellNumberText].Value);

                    cross_descr.AppendChild(cross_descr_text);
                    cross.Attributes.Append(cross_descr);
                    cross.AppendChild(cross_text);
                    crossess.AppendChild(cross);
                    cross_text = null;
                    cross_descr_text = null;
                    cross_descr = null;
                    cross = null;

                }

                userElem.Attributes.Append(slack_id);
                userElem.Attributes.Append(to_fit);
                userElem.Attributes.Append(axle_side);
                userElem.Attributes.Append(other_side);
                userElem.Attributes.Append(vehicle_type);
                userElem.Attributes.Append(slack_type);
                userElem.Attributes.Append(trand_mark);

                //crossess.
                userElem.AppendChild(crossess);
                userElem.AppendChild(hole_params);
                userElem.AppendChild(slack_params);
                xRoot.AppendChild(userElem);

                richTextBox1.AppendText(slack_id.Value + " | " + type_mark.SlactType + " | " + type_mark.TrandMark + "\r\n");
                //Зачистка переменных

                slack_id = null;
                to_fit = null;
                axle_side = null;
                other_side = null;
                vehicle_type = null;
                crossess = null;
                hole_params = null;
                slack_params = null;
                slack_type = null;
                trand_mark = null;

                slack_id_text = to_fit_text = axle_side_text = other_side_text = vehicle_type_text = slack_type_text = trand_mark_text = null;
                userElem = null;


            }
            //добавляем узлы
            xDoc.Save("slk2.xml");
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            /**
             * Загружаем данные из старого кроссного файла
             */
            //Формирование списка рычагов на основании файла EAC Certification List June 2018
            //Открываем диалог.
            //Если отказываем - выход
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            //Присвоение имени файла
            string fileName = openFileDialog1.FileName;

            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();

            excelApp.Visible = true;
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(fileName);
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];

            /*cSlack slk = new cSlack();
            string val = wrkSheet.Range["A4"].Value;
            slk.id = val.ToString();

            richTextBox1.AppendText(slk.id);*/

            string[] input = { "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };
            List<string> colls = new List<string>(input);
            string cellNumberText, cellNumberValue;

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

        }

        private void xmlToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //Открываем диалог сохранения
            if (sFD.ShowDialog() == DialogResult.Cancel)
                return;
            saveSlacks(sFD.FileName);
        }

        private void saveSlacks(String fileName)
        {
            
            //Сохранение данных из списка рычагов в уквазанном файле xml
            XmlDocument xDoc = new XmlDocument();

            xDoc.Load("slacks.xml");
            XmlElement xRoot = xDoc.DocumentElement;

            foreach (cSlack slk in this.slacks)
            {
                xRoot.AppendChild(slk.ToXmlElement(xDoc));
            }
            xDoc.Save(fileName);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            this.tbCrosses.Text = openFileDialog1.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            this.tbCataloge.Text = openFileDialog1.FileName;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            this.tbBarcodes.Text = openFileDialog1.FileName;
        }

        private void bntCrosses_Click(object sender, EventArgs e)
        {
            //Заполняем элементы массива Slacks данными из файла с кросс номерами.
            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();
            //Видимость открытого экземпляра файла
            excelApp.Visible = true;
            //Открываем книгу
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(this.tbCrosses.Text);
            //Открываем закладку
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];
            //Массив наименований колонок
            List<string> colls = new List<string> { "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L" };
            //Для каждого элемента массива slacks проводим процедуру поиска
            foreach(cSlack slk in this.slacks)
            {
                Excel.Range currentRange, cRg;
                //Получаем диапазон клеток
                cRg = wrkSheet.Columns["A:A"];
                //Запускаем поиск
                currentRange = cRg.Find(slk.id);

                if (currentRange != null)   //Артикул найден
                {
                    richTextBox1.AppendText("Артикул " + slk.id + " найден в ячейке " + currentRange.Row.ToString() + ". Загрузка данных в объект\r\n");
                    //Тип транспортного средства
                    if (wrkSheet.Range["N" + currentRange.Row].Value != null)
                        slk.vehicleType = wrkSheet.Range["N" + currentRange.Row].Value;
                    //Кроссы - crossess
                    for (int i = 0; i < colls.Count(); i++)
                    {

                        switch (wrkSheet.Range[colls[i] + "1"].Value)
                        {
                            //Кроссовый номер для QAS
                            case "QAS":
                                if (wrkSheet.Range[colls[i] + currentRange.Row].Value != null)
                                    slk.qasId = wrkSheet.Range[colls[i] + currentRange.Row].Value;
                                break;

                            //Кроссовый номер для HALDEX
                            case "HAL1":
                            case "HAL2":
                            case "HAL3":
                            case "HAL4":
                            case "HAL5":
                                if (wrkSheet.Range[colls[i] + currentRange.Row].Value != null)
                                    slk.haldexId.Add(wrkSheet.Range[colls[i] + currentRange.Row].Value.ToString());
                                break;

                            //Кроссовый номер для OEM
                            case "OE1":
                            case "OE2":
                            case "OE3":
                            case "OE4":
                            case "OE5":
                                if (wrkSheet.Range[colls[i] + currentRange.Row].Value != null)
                                    slk.oemId.Add(wrkSheet.Range[colls[i] + currentRange.Row].Value.ToString());
                                break;
                        }

                    }

                }
                else
                {
                    richTextBox2.AppendText("Артикул " + slk.id + " НЕ найден\r\n");
                }
                //richTextBox1.AppendText(slack_id.Value + " | " + type_mark.SlactType + " | " + type_mark.TrandMark + "\r\n");
            }
        }

        private void bntCatalog_Click(object sender, EventArgs e)
        {
            //Загружаем общие данные из каталога

            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();

            excelApp.Visible = true;
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(this.tbCataloge.Text);
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];

            /**
             * Запускаем поиск по подгруженному файлу
             */
            int i = 0, tmp090 = 0, rowIndex = 0;
            String val;
            foreach (cSlack slk in this.slacks)
            {

                Excel.Range currentRange, cRg;

                cRg = wrkSheet.Columns["A:A"];
                currentRange = cRg.Find(slk.id);//, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext
                //

                if (currentRange != null)
                {
                    richTextBox1.AppendText("Артикул " + slk.id + " найден в ячейке " + currentRange.Row.ToString() + ". Загрузка данных в объект\r\n");
                    //Короссы Haldex
                    rowIndex = currentRange.Row + 1;
                    val = wrkSheet.Range["K" + rowIndex.ToString()].Value.Trim('H', 'a', 'l','d','e','x', ':', ' ');
                    if (slk.haldexId.IndexOf(val) < 0)
                    {
                        slk.haldexId.Add(val);
                    }

                    //Короссы OEM
                    rowIndex = currentRange.Row + 2;
                    val = wrkSheet.Range["K" + rowIndex.ToString()].Value.Trim('O', 'E', 'M', ':', ' ');

                    if(val.Count() > 4 && slk.oemId.IndexOf(val) < 0)
                    {
                        slk.oemId.Add(val);
                    }
                    

                    //Комплект поставки
                    rowIndex = currentRange.Row + 3;
                    val = wrkSheet.Range["K" + rowIndex.ToString()].Value;
                    slk.suppliedWith = val.Substring(14, val.Length - 14).TrimStart(' ');
                    //Сторона
                    rowIndex = currentRange.Row + 4;
                    val = wrkSheet.Range["K" + rowIndex.ToString()].Value;
                    slk.axleSide = val.Substring(9, val.Length - 9).TrimStart(':');
                    //Другая сторона
                    val = wrkSheet.Range["T" + rowIndex.ToString()].Value;
                    slk.otherSide = val.Substring(11, val.Length - 11).TrimStart(' ');

                    rowIndex = currentRange.Row + 6;
                    //Смещение
                    val = (wrkSheet.Range["K" + rowIndex.ToString()].Value != null ? wrkSheet.Range["K" + rowIndex.ToString()].Value.ToString() : "0");
                    Int32.TryParse(val.Trim(), out tmp090);
                    slk.offset = tmp090;


                    //Наклон
                    //            double inc = wrkSheet.Range["N7"].Value;
                    val = (wrkSheet.Range["N" + rowIndex.ToString()].Value != null ? wrkSheet.Range["N" + rowIndex.ToString()].Value.ToString() : "0");
                    Int32.TryParse(val.Trim(), out tmp090);
                    slk.inclination = tmp090;

                    //Угол поводка
                    val = wrkSheet.Range["R" + rowIndex.ToString()].Value != null ? wrkSheet.Range["R" + rowIndex.ToString()].Value.ToString() : "";
                    slk.controlArmAngle = val.Trim();
                    //Int32.TryParse(val.Trim(), out cSlack.controlArmAngle);

                    //Тип поводка: AP, QF
                    val = wrkSheet.Range["V" + rowIndex.ToString()].Value != null ? wrkSheet.Range["V" + rowIndex.ToString()].Value.ToString() : "";
                    slk.controlArmType = val.Trim(); ;

                    //Количество зубьев
                    val = wrkSheet.Range["Y" + rowIndex.ToString()].Value != null ? wrkSheet.Range["Y" + rowIndex.ToString()].Value.ToString() : "0";

                    Int32.TryParse(val.Trim(), out tmp090);

                    slk.splineTeeth = tmp090;

                    //Вилочная втулка: 14.2,
                    val = wrkSheet.Range["AC" + rowIndex.ToString()].Value.ToString();
                    slk.clevisBush = val.Trim();

                    //Расстояние между отверстиями
                    int tmp1 = 0, tmp2 = 0, tmp3 = 0, tmp4 = 0, tmp5 = 0, tmp6 = 0, tmp7 = 0;
                    rowIndex = currentRange.Row + 8;
                    Int32.TryParse((wrkSheet.Range["K" + rowIndex.ToString()].Value != null ? wrkSheet.Range["K" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp1);
                    Int32.TryParse((wrkSheet.Range["N" + rowIndex.ToString()].Value != null ? wrkSheet.Range["N" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp2);
                    Int32.TryParse((wrkSheet.Range["Q" + rowIndex.ToString()].Value != null ? wrkSheet.Range["Q" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp3);
                    Int32.TryParse((wrkSheet.Range["T" + rowIndex.ToString()].Value != null ? wrkSheet.Range["T" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp4);
                    Int32.TryParse((wrkSheet.Range["W" + rowIndex.ToString()].Value != null ? wrkSheet.Range["W" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp5);

                    //string str1 = wrkSheet.Range["Z9"].Value;
                    Int32.TryParse((wrkSheet.Range["Z" + rowIndex.ToString()].Value != null ? wrkSheet.Range["Z" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp6);

                    //str1 = wrkSheet.Range["AC9"].Value;
                    Int32.TryParse((wrkSheet.Range["AC" + rowIndex.ToString()].Value != null ? wrkSheet.Range["AC" + rowIndex.ToString()].Value.ToString().Trim() : "0"), out tmp7);
                    //val = wrkSheet.Range["T5"].Value;
                    slk.holeCentreSize = new List<int> { tmp1, tmp2, tmp3, tmp4, tmp5, tmp6, tmp7 };
                }
                else
                {
                    richTextBox2.AppendText("Ошибка: Артикул " + slk.id + " НЕ НАЙДЕН!\r\n");
                }
            }
        }

        private void btnBarcodes_Click(object sender, EventArgs e)
        {
            /**
             * Подгружаем в данные по весу, габаритам, штрихкоду и упаковке
             */
            //Открываем файл и грузим в Excel.
            Excel.Application excelApp = new Excel.Application();

            excelApp.Visible = true;
            Excel.Workbook wrkBook = excelApp.Workbooks.Open(this.tbBarcodes.Text);
            Excel.Worksheet wrkSheet = (Excel.Worksheet)wrkBook.Worksheets[1];

            /**
             * Запускаем поиск по подгруженному файлу
             */
            //int rowIndex = 0; //i = 0
            String val;
            foreach (cSlack slk in this.slacks)
            {

                Excel.Range currentRange, cRg;

                cRg = wrkSheet.Columns["A:A"];
                currentRange = cRg.Find(slk.id);//, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext
                //

                if (currentRange != null)
                {
                    richTextBox1.AppendText("Артикул " + slk.id + " найден в ячейке " + currentRange.Row.ToString() + ". Загрузка данных в объект\r\n");
                    
                    //Штрихкод
                    val = wrkSheet.Range["B" + currentRange.Row.ToString()].Value.ToString();
                    if (val != null)
                        slk.barcode = val;

                    //Вес
                    double tmp = wrkSheet.Range["C" + currentRange.Row.ToString()].Value;
                    val = tmp.ToString();
                    if (val != null)
                        slk.weight =Convert.ToDouble(val);
                    
                    //Тип упаковки
                    val = wrkSheet.Range["D" + currentRange.Row.ToString()].Value.ToString();
                    if(val != null)
                        slk.boxType = Convert.ToInt32(val);
                    
                    //Размеры
                    val = wrkSheet.Range["E" + currentRange.Row.ToString()].Value.ToString();
                    slk.dimension = val;
                }
                else
                {
                    richTextBox2.AppendText("Ошибка: Артикул " + slk.id + " НЕ НАЙДЕН!\r\n");
                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (folderBrowserDialog1.ShowDialog() == DialogResult.Cancel)
                return;

            
            this.tbImages.Text = folderBrowserDialog1.SelectedPath;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Подгружаем картинки в xml файл
            string filePrefix = "_mn";
            DirectoryInfo dir = new DirectoryInfo(this.tbImages.Text);
            //Перебираем все объекты
            //FileInfo[] filesInDir = dir.GetFiles("*.jpg", SearchOption.AllDirectories))
            bool FileExists = false;
            string startupPath = System.IO.Directory.GetCurrentDirectory();

            foreach (cSlack slk in this.slacks)
            {
                foreach (var fInfo in dir.GetFiles("*" + slk.id.ToString() + "*.jpg", SearchOption.AllDirectories))
                {
                    foreach(cSlackImages slkImg in slk.imagesNames)
                    {
                        if (slkImg.FileCompare(fInfo.FullName))
                            FileExists = true;
                    }
                    if (!FileExists)
                    {
                        //Создаём экземпляр класса
                        cSlackImages cSlackImage = new cSlackImages();
                        //Копируем файл в целевую папку и присваиваем имя на основании артикула и количества файлов
                        string newFilePostfix = (slk.imagesNames.Count > 0 ? "_"+slk.imagesNames.Count.ToString() : "");
                        DirectoryInfo newDirInfo = new DirectoryInfo(startupPath + "\\images\\");
                        FileInfo newFile;
                        if (!File.Exists(startupPath + "\\images\\" + slk.id + newFilePostfix + ".jpg")) { 
                            newFile = fInfo.CopyTo(startupPath+"\\images\\"+slk.id+newFilePostfix+".jpg");
                        }
                        else
                        {
                            newFile = newDirInfo.GetFiles(startupPath + "\\images\\" + slk.id + newFilePostfix + ".jpg")[0];
                        }
                        cSlackImage.fName   = slk.id + newFilePostfix + ".jpg";
                        cSlackImage.fSize = newFile.Length;
                        cSlackImage.imagePath = startupPath + "\\images\\";
                        slk.imagesNames.Add(cSlackImage);
                    }
                }
                //foreach (cSlackImages slkImg in filesInDir.Length){}
                //if(slk.id.ToString()+".jp")
                //Ищем для каждого последующего объекта подходящий файл
                /*if (slk.imagesNames.Count() > 0){foreach ( cSlackImages slkImage in slk.imagesNames){if(slkImage.fSize == fInfo.Length){}}}}*/

            }
        }
    }
}
