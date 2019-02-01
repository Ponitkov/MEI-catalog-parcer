using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace excelParcer
{
    class cSlack
    {
        //Slack ID/Article
        public string id { get; set;  }

        //Slack to fit
        public string toFit { get; set; }

        //MEI/QAS
        public string trendMark { get; set; }

        public string vehicleType { get; set; }

        //Slack type can be: for MEI MSA/ASA/S-ASA, for QAS QAS/SS
        public string slackType { get; set; }
        //Аналог QAS 
        public string qasId { get; set; }
        //Haldex ID array
        public List<string> haldexId;

        //OEM ID array
        public List<string> oemId;

        //Supplied with
        public string suppliedWith { get; set; }

        // Axle side can be: L, R, L&R
        public string axleSide { get; set; }

        //Other hand
        public string otherSide { get; set; }

        //Offset can be - or +
        public int offset { get; set; }

        //Inclination can be - or +
        public int inclination { get; set; }

        //Control arm angle 
        public string controlArmAngle { get; set; }

        //Control arm type can be like: AP, QF
        public string controlArmType { get; set; }

        //Spline teeth can be: 10, 12, 14, >20
        public int splineTeeth { get; set; }

        //Clevis buch can be: 14.2,
        public string clevisBush { get; set; }
        
        //Вес
        public double weight { get; set; }
        //Тип упаковки
        public int boxType { get; set; }
        //Штрихкод
        public string barcode { get; set; }
        //Размер
        public string dimension { get; set; }

        //Hole centre can have max 7 values
        public List<int> holeCentreSize;

        //Images
        public List<cSlackImages> imagesNames;

        //Конструктор класса
        public cSlack()
        {
            this.haldexId = new List<string>();
            this.oemId = new List<string>();
            this.holeCentreSize = new List<int>();
            this.imagesNames = new List<cSlackImages>();
        }

        public bool LoadFrom(XmlNode elm)
        {
            //При передаче пустого элемента - возвращаем отказ
            if(elm == null)
                return false;

            this.id = elm.Attributes["id"].Value.ToString(); ;
            this.slackType = elm.Attributes["slack_type"].Value.ToString();
            this.toFit = elm.Attributes["to_fit"].Value.ToString();
            this.axleSide = elm.Attributes["axle_side"].Value.ToString();
            this.otherSide = elm.Attributes["other_side"].Value.ToString();
            this.vehicleType = elm.Attributes["vehicle_type"].Value.ToString();
            this.trendMark = elm.Attributes["trand_mark"].Value.ToString();
            
            //Получаем узел с кроссами
            XmlNode crosses = elm.FirstChild;

            foreach (XmlNode cross in crosses.ChildNodes)
            {
                switch (cross.Attributes["descr"].Value.ToString())
                {
                    case "QAS":
                            if (cross.InnerText != "")
                                this.qasId = cross.InnerText;
                        break;

                    case "HAL1":
                    case "HAL2":
                    case "HAL3":
                    case "HAL4":
                    case "HAL5":
                            if (cross.InnerText != "")
                                this.haldexId.Add(cross.InnerText);
                        break;

                    case "OE1":
                    case "OE2":
                    case "OE3":
                    case "OE4":
                    case "OE5":
                            if (cross.InnerText != "")
                                this.oemId.Add(cross.InnerText);
                        break;
                }
            }

            //Получаем узел с кроссами
            XmlNode holes = crosses.NextSibling;
            foreach(XmlNode hole in holes.ChildNodes)
            {
                this.holeCentreSize.Add(Convert.ToInt32(hole.InnerText));
            }

            //Подгружаем остальные параметры
            XmlNode slackParams = holes.NextSibling;
            for (int  i = 0; i < slackParams.ChildNodes.Count; i++)
            {
                switch (slackParams.ChildNodes[i].LocalName)
                {
                    //Смещение
                    case "offset":
                        this.offset = Convert.ToInt32(slackParams.ChildNodes[i].InnerText);
                        break;

                    //Наклон
                    case "inclination":
                        this.inclination = Convert.ToInt32(slackParams.ChildNodes[i].InnerText);
                        break;

                    //Угол поводка
                    case "control_arm_angle":
                        this.controlArmAngle = slackParams.ChildNodes[i].InnerText;
                        break;

                    //Тип поводка
                    case "control_arm_type":
                        this.controlArmType = slackParams.ChildNodes[i].InnerText;
                        break;

                    //Количество зубьев
                    case "spline_teeth":
                        this.splineTeeth = Convert.ToInt32(slackParams.ChildNodes[i].InnerText);
                        break;
                    
                    //Вилочная втулка
                    case "clevis_bush":
                        this.clevisBush = slackParams.ChildNodes[i].InnerText;
                        break;
                    
                    //Вес
                    case "weight":
                        this.weight = Convert.ToDouble(slackParams.ChildNodes[i].InnerText);
                        break;
                    
                    //Тип упаковки
                    case "box_type":
                        this.boxType = Convert.ToInt32(slackParams.ChildNodes[i].InnerText);
                        break;

                    //Штрихкод
                    case "barcode":
                        this.barcode = slackParams.ChildNodes[i].InnerText;
                        break;
                    
                    //Размер
                    case "dimension":
                        this.dimension  = slackParams.ChildNodes[i].InnerText;
                        break;
                }
            }

            //Подгружаем картинки
            XmlNode slackImages = slackParams.NextSibling;
            try
            {
                foreach (XmlNode image in slackImages.ChildNodes)
                {
                    cSlackImages slkImg = new cSlackImages();
                    slkImg.FromXMLElement(image);
                    this.imagesNames.Add(slkImg);
                }
            }catch(Exception e)
            {
                ;
            }

            return true;

        }

        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override string ToString()
        {
            return base.ToString();
        }
        
        public XmlElement ToXmlElement(XmlDocument slkDoc)
        {
            XmlElement elem = slkDoc.CreateElement("slack");
            // создаем атрибут id
            XmlAttribute slack_id = slkDoc.CreateAttribute("id");
            // создаем атрибут to_fit
            XmlAttribute to_fit = slkDoc.CreateAttribute("to_fit");
            // создаем атрибут axle_side
            XmlAttribute axle_side = slkDoc.CreateAttribute("axle_side");
            // создаем атрибут other_side
            XmlAttribute other_side = slkDoc.CreateAttribute("other_side");
            // создаем атрибут vehicle_type
            XmlAttribute vehicle_type = slkDoc.CreateAttribute("vehicle_type");
            // создаем атрибут type
            XmlAttribute slack_type = slkDoc.CreateAttribute("slack_type");
            // создаем атрибут trand_mark
            XmlAttribute trand_mark = slkDoc.CreateAttribute("trand_mark");


            // создаем элементы crossess
            XmlElement crossess = slkDoc.CreateElement("crossess");
            // создаём элемент hole_params
            XmlElement hole_params = slkDoc.CreateElement("hole_params");
            // создаём элемент slack_params
            XmlElement slack_params = slkDoc.CreateElement("slack_params");

            // создаем текстовые значения для элементов и атрибута
            // артикул
            XmlText slack_id_text = slkDoc.CreateTextNode(this.id);

            //CSlackType type_mark = new CSlackType(wrkSheet.Range["A" + j].Value);
            XmlText slack_type_text = slkDoc.CreateTextNode(this.slackType);
            XmlText trand_mark_text = slkDoc.CreateTextNode(this.trendMark);
            // марка транспортного средства/производитель оси
            XmlText to_fit_text = slkDoc.CreateTextNode(toFit);
            // сторона
            XmlText axle_side_text = slkDoc.CreateTextNode(this.axleSide);
            // другая сторона
            XmlText other_side_text = slkDoc.CreateTextNode(this.otherSide);
            // тип транспортного средства
            XmlText vehicle_type_text = slkDoc.CreateTextNode(this.vehicleType);

            //Аттрибуты
            slack_id.AppendChild(slack_id_text);
            to_fit.AppendChild(to_fit_text);
            axle_side.AppendChild(axle_side_text);
            other_side.AppendChild(other_side_text);
            vehicle_type.AppendChild(vehicle_type_text);
            slack_type.AppendChild(slack_type_text);
            trand_mark.AppendChild(trand_mark_text);

            //Кроссы - crossess
            XmlElement cross;
            XmlAttribute cross_descr;
            XmlText cross_descr_text, cross_text;
            //Сначала Haldex
            foreach (string hldId in this.haldexId)
            {
                cross = slkDoc.CreateElement("cross");
                cross_descr = slkDoc.CreateAttribute("descr");
                cross_descr_text = slkDoc.CreateTextNode("Haldex");
                cross_text = slkDoc.CreateTextNode(hldId);

                cross_descr.AppendChild(cross_descr_text);
                cross.Attributes.Append(cross_descr);
                cross.AppendChild(cross_text);
                crossess.AppendChild(cross);
                cross_text = null;
                cross_descr_text = null;
                cross_descr = null;
                cross = null;
            }
            //Теперь OE
            foreach (string OEId in this.oemId)
            {
                cross = slkDoc.CreateElement("cross");
                cross_descr = slkDoc.CreateAttribute("descr");
                cross_descr_text = slkDoc.CreateTextNode("OEM");
                cross_text = slkDoc.CreateTextNode(OEId);

                cross_descr.AppendChild(cross_descr_text);
                cross.Attributes.Append(cross_descr);
                cross.AppendChild(cross_text);
                crossess.AppendChild(cross);
                cross_text = null;
                cross_descr_text = null;
                cross_descr = null;
                cross = null;
            }

            if (this.qasId != null && this.qasId != "")
            {
                cross = slkDoc.CreateElement("cross");
                cross_descr = slkDoc.CreateAttribute("descr");
                cross_descr_text = slkDoc.CreateTextNode("QAS");
                cross_text = slkDoc.CreateTextNode(this.qasId);

                cross_descr.AppendChild(cross_descr_text);
                cross.Attributes.Append(cross_descr);
                cross.AppendChild(cross_text);
                crossess.AppendChild(cross);
            }

            elem.Attributes.Append(slack_id);
            elem.Attributes.Append(to_fit);
            elem.Attributes.Append(axle_side);
            elem.Attributes.Append(other_side);
            elem.Attributes.Append(vehicle_type);
            elem.Attributes.Append(slack_type);
            elem.Attributes.Append(trand_mark);

            //crossess
            elem.AppendChild(crossess);
            //Расстояния по отверстиям
            XmlElement hole_centre;
            XmlAttribute hole_centre_descr;
            XmlText hole_centre_descr_text, hole_centre_text;
            for (int i = 0; i < this.holeCentreSize.Count(); i++)
            {
                hole_centre = slkDoc.CreateElement("centre");
                hole_centre_descr = slkDoc.CreateAttribute("position");
                hole_centre_descr_text = slkDoc.CreateTextNode(i.ToString());
                hole_centre_text = slkDoc.CreateTextNode(this.holeCentreSize[i].ToString());

                hole_centre_descr.AppendChild(hole_centre_descr_text);
                hole_centre.Attributes.Append(hole_centre_descr);
                hole_centre.AppendChild(hole_centre_text);
                hole_params.AppendChild(hole_centre);
                hole_centre_descr = null;
                hole_centre_descr_text = null;
                hole_centre = null;
                hole_centre_text = null;
            }
            elem.AppendChild(hole_params);

            //Параметры рычага.
            //Наименование: Смещение
            XmlElement offset = slkDoc.CreateElement("offset");
            //Значение
            XmlText offset_text = slkDoc.CreateTextNode(this.offset.ToString());
            offset.AppendChild(offset_text);
            //Наименование: Наклон
            XmlElement inclination = slkDoc.CreateElement("inclination");
            //Значение
            XmlText inclination_text = slkDoc.CreateTextNode(this.inclination.ToString());
            inclination.AppendChild(inclination_text);
            //Наименование: Угол поводка
            XmlElement control_arm_angle = slkDoc.CreateElement("control_arm_angle");
            //Значение
            XmlText control_arm_angle_text = slkDoc.CreateTextNode(this.controlArmAngle == null ? "" : this.controlArmAngle.ToString());
            control_arm_angle.AppendChild(control_arm_angle_text);
            //Наименование: Тип поводка
            XmlElement controlArmType = slkDoc.CreateElement("control_arm_type");
            //Значение
            XmlText controlArmType_text = slkDoc.CreateTextNode(this.controlArmType == null ? "" : this.controlArmType.ToString());
            controlArmType.AppendChild(controlArmType_text);
            //Наименование: Количество зубьев
            XmlElement splineTeeth = slkDoc.CreateElement("spline_teeth");
            //Значение
            XmlText splineTeeth_text = slkDoc.CreateTextNode(this.splineTeeth.ToString());
            splineTeeth.AppendChild(splineTeeth_text);
            //Наименование: Вилочная втулка
            XmlElement clevisBush = slkDoc.CreateElement("clevis_bush");
            //Значение
            XmlText clevisBush_text = slkDoc.CreateTextNode(this.clevisBush == null ? "" : this.clevisBush.ToString());
            clevisBush.AppendChild(clevisBush_text);

            //Наименование: Вес
            XmlElement weight = slkDoc.CreateElement("weight");
            //Значение
            XmlText weight_text = slkDoc.CreateTextNode(this.weight.ToString());
            weight.AppendChild(weight_text);

            //Наименование: Тип упаковки
            XmlElement boxType = slkDoc.CreateElement("box_type");
            //Значение
            XmlText boxType_text = slkDoc.CreateTextNode(this.boxType.ToString());
            boxType.AppendChild(boxType_text);

            //Наименование: Штрихкод
            XmlElement barcode = slkDoc.CreateElement("barcode");
            //Значение
            XmlText barcode_text = slkDoc.CreateTextNode(this.barcode == null ? "" : this.barcode.ToString());
            barcode.AppendChild(barcode_text);

            //Наименование: Размер
            XmlElement dimension = slkDoc.CreateElement("dimension");
            //Значение
            XmlText dimension_text = slkDoc.CreateTextNode(this.dimension == null ? "" : this.dimension.ToString());
            dimension.AppendChild(dimension_text);


            slack_params.AppendChild(offset);
            slack_params.AppendChild(inclination);
            slack_params.AppendChild(control_arm_angle);
            slack_params.AppendChild(controlArmType);
            slack_params.AppendChild(splineTeeth);
            slack_params.AppendChild(clevisBush);
            slack_params.AppendChild(weight);
            slack_params.AppendChild(boxType);
            slack_params.AppendChild(barcode);
            slack_params.AppendChild(dimension);

            elem.AppendChild(slack_params);
            //images
            XmlElement slackImages = slkDoc.CreateElement("slack_images");
            foreach (cSlackImages slkImage in this.imagesNames)
            {
                slackImages.AppendChild(slkImage.ToXMLNode(slkDoc));
            }

            elem.AppendChild(slackImages);

            return elem;
        }
    }
}
