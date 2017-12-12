using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Reflection;
using Newtonsoft.Json.Linq;

namespace ExprotData
{
    class ExportJSON
    {
        int year, month, day, hour, minute, second;
        public string ExportJSON_Method(string CsvPath, string Filename)
        {
          
            Filename = Filename.Replace(".csv", "");

            //var mapStr = Server.MapPath("~/Excel文件夹/") + "724" + ".xls";
            var mapStr = CsvPath;
            Form1.form1.progressBar2.Value = 0;
            Form1.form1.label2.Text = "正在打开文件" + Filename;
            var LstData = AsposeTool.ReadExcel(mapStr);//GetJSonData(DTStart, DTEnd, VIN);
            List<string> JSList = new List<string>();

            string key = "";
            string value = "";
            string unitnumber = "";
            if (LstData.Count > 0)
            {
                DataTable dT = new DataTable();
                dT.Columns.Add("时间");
                Form1.form1.label2.Text = "正在分析" + "文件" + Filename + "的数据";
                Console.Write(DateTime.Now+"正在分析"+"文件"+Filename + "的数据-");
                Form1.form1.progressBar2.Maximum =  LstData.Count;
                
                for (int i = 0; i < LstData.Count; i++)
                {
                    Dictionary<string, string> Key_Value = new Dictionary<string, string>();
                    #region 方法2
                    //var o = JObject.Parse(LstData[i]);

                    //foreach (JToken child in o.Children())
                    //{
                    //    //var property1 = child as JProperty;  
                    //    //MessageBox.Show(property1.Name + ":" + property1.Value);  
                    //    foreach (JToken grandChild in child)
                    //    {
                    //        var property = child as JProperty;
                    //        foreach (JToken grandGrandChild in grandChild)
                    //        {
                    //             property = grandGrandChild as JProperty;
                    //            //if (grandGrandChild.Name.Equals("整车"))
                    //            //{
                    //            //    foreach ()


                    //            //}
                    //            foreach (JToken grandGrandChild1 in grandGrandChild) {

                    //                property = grandGrandChild as JProperty;
                    //                //if (property != null)
                    //                //{
                    //                //    Console.WriteLine(property.Name + ":" + property.Value);
                    //                //}



                    //            }





                    //            if (property != null)
                    //            {
                    //                Console.WriteLine(property.Name + ":" + property.Value);
                    //            }
                    //        }
                    //        if (property != null)
                    //        {
                    //            Console.WriteLine(property.Name + ":" + property.Value);
                    //        }
                    //    }
                    //}






                    #endregion




                    #region 方法1


                    RootObject rb = JsonConvert.DeserializeObject<RootObject>(LstData[i]);
                    Type t = rb.GetType();//获得该类的Type
                    foreach (PropertyInfo pi in t.GetProperties()) {
                        object value1 = pi.GetValue(rb, null);   //用pi.GetValue获得值
                        Type t1 = value1.GetType();//获得该类的Type
                        if (pi.PropertyType == typeof(string))//属性的类型判断  
                        {
                            //object value1 = pi.GetValue(rb, null);   //用pi.GetValue获得值
                            //Type t1 = value1.GetType();//获得该类的Type
                            //object obj = i.GetValue(contract, null);
                            string name1 = pi.Name;//获得属性的名字,后面就可以根据名字判断来进行些自己想要的操作
                                                   //进行你想要的操作
                            //Console.WriteLine(name1 + "\t" + value1);
                            Key_Value.Add(name1, (string)value1);
                            continue;
                        }


                        else if (pi.PropertyType == typeof(Fault)) {
                            foreach (PropertyInfo pi2 in t1.GetProperties()) {
                                if (pi2.PropertyType == typeof(List<故障>))
                                {
                                    List<故障> vlist = rb.fault.故障;
                                    if (vlist != null)
                                    {
                                        for (int j = 0; j < vlist.Count; j++)
                                        {
                                            //vlist[i].description;

                                            //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                            key = vlist[j].description;
                                            value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                            Key_Value.Add(key, value);


                                        }


                                    }

                                }else if (pi2.PropertyType == typeof(List<报警>))
                                {
                                    List<报警> vlist = rb.fault.报警;
                                    if (vlist != null)
                                    {
                                        for (int j = 0; j < vlist.Count; j++)
                                        {
                                            //vlist[i].description;

                                            //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                            key = vlist[j].description;
                                            value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                            Key_Value.Add(key, value);


                                        }


                                    }

                                }


                            }

                            continue;
                        }
                        else if (pi.PropertyType == typeof(List<Voltage>))//属性的类型判断  
                        {
                            //object value1 = pi.GetValue(rb, null);   //用pi.GetValue获得值
                            List<Voltage> vlist = rb.voltage;
                            for (int j = 0; j < vlist.Count; j++)
                            {
                                //vlist[i].description;

                                //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                key = vlist[j].description;
                                value = vlist[j].value;
                                Key_Value.Add(key, value);

                            }
                            //object obj = i.GetValue(contract, null);

                            //string name1 = pi.Name;//获得属性的名字,后面就可以根据名字判断来进行些自己想要的操作
                            //                       //进行你想要的操作
                            //Console.WriteLine(name1 + "\t" + value1);
                            continue;
                        }
                        else if (pi.PropertyType == typeof(List<Temperature>))//属性的类型判断  
                        {
                            //object value1 = pi.GetValue(rb, null);   //用pi.GetValue获得值
                            List<Temperature> vlist = rb.temperature;
                            for (int j = 0; j < vlist.Count; j++)
                            {
                                //vlist[i].description;

                                //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                key = vlist[j].description;
                                value = vlist[j].value;
                                Key_Value.Add(key, value);

                            }
                            //object obj = i.GetValue(contract, null);

                            //string name1 = pi.Name;//获得属性的名字,后面就可以根据名字判断来进行些自己想要的操作
                            //                       //进行你想要的操作
                            //Console.WriteLine(name1 + "\t" + value1);
                            continue;
                        }

                        foreach (PropertyInfo pi2 in t1.GetProperties())
                        {
                            if (pi2.PropertyType == typeof(List<整车>))
                            {
                                List<整车> vlist = rb.vehicle.整车;
                                for (int j = 0; j < vlist.Count; j++)
                                {
                                    //vlist[i].description;

                                    //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                    key = vlist[j].description;
                                    value = value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                    Key_Value.Add(key, value);

                                }
                                continue;
                            }
                            else if (pi2.PropertyType == typeof(List<故障>))
                            {
                                List<故障> vlist = rb.vehicle.故障;
                                if (vlist != null) {
                                    for (int j = 0; j < vlist.Count; j++)
                                    {
                                        //vlist[i].description;

                                        //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                        key = vlist[j].description;
                                        value = value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                        Key_Value.Add(key, value);

                                    }


                                }
                                continue;
                            }
                            else if (pi2.PropertyType == typeof(List<驱动>))
                            {
                                List<驱动> vlist = rb.vehicle.驱动;
                                if (vlist != null)
                                {
                                    for (int j = 0; j < vlist.Count; j++)
                                    {
                                        //vlist[i].description;

                                        //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                        key = vlist[j].description;
                                        value = value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                        Key_Value.Add(key, value);

                                    }
                                }
                                continue;
                            }
                            else if (pi2.PropertyType == typeof(List<发动机>))
                            {
                                List<发动机> vlist = rb.vehicle.发动机;
                                if (vlist != null)
                                {
                                    for (int j = 0; j < vlist.Count; j++)
                                    {
                                        //vlist[i].description;

                                        //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                        key = vlist[j].description;
                                        value = value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                        Key_Value.Add(key, value);

                                    }
                                }
                                continue;
                            }
                            else if (pi2.PropertyType == typeof(List<燃料电池>))
                            {
                                List<燃料电池> vlist = rb.vehicle.燃料电池;
                                if (vlist != null)
                                {
                                    for (int j = 0; j < vlist.Count; j++)
                                    {
                                        //vlist[i].description;

                                        //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                        key = vlist[j].description;
                                        value = value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                        Key_Value.Add(key, value);

                                    }
                                }
                                continue;
                            }
                            else if (pi2.PropertyType == typeof(List<极值>))
                            {
                                List<极值> vlist = rb.vehicle.极值;
                                if (vlist != null)
                                {
                                    for (int j = 0; j < vlist.Count; j++)
                                    {
                                        //vlist[i].description;

                                        //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                        key = vlist[j].description;
                                        value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                        Key_Value.Add(key, value);

                                    }
                                }
                                continue;
                            }
                            else if (pi2.PropertyType == typeof(List<报警>))
                            {
                                List<报警> vlist = rb.vehicle.报警;
                                if (vlist != null)
                                {
                                    for (int j = 0; j < vlist.Count; j++)
                                    {
                                        //vlist[i].description;

                                        //Console.WriteLine(vlist[j].description + "\t" + vlist[j].value);
                                        key = vlist[j].description;
                                        value = chooseResult(vlist[j].value, vlist[j].valueTable);
                                        Key_Value.Add(key, value);

                                    }
                                }
                                continue;
                            }

                            object value2 = pi2.GetValue(value1, null);   //用pi.GetValue获得值
                            string name2 = pi2.Name;//获得属性的名字,后面就可以根据名字判断来进行些自己想要的操作
                                                    //进行你想要的操作
                            //Console.WriteLine(name2 + "\t" + value2);
                            key = name2;
                            value = (string)value2;
                            Key_Value.Add(key, value);
                        }



                        //if (!t1.Name.Equals("String")) {

                        //    foreach (PropertyInfo pi1 in t1.GetProperties())
                        //    {
                        //        object value2 = pi1.GetValue(t1, null);   //用pi.GetValue获得值
                        //        string name2 = pi1.Name;//获得属性的名字,后面就可以根据名字判断来进行些自己想要的操作
                        //                                //进行你想要的操作
                        //        Console.WriteLine(name2 + "\t"+value2);
                        //    }
                        //}
                        //else {





                        //}
                    }


                    #endregion


                    #region 旧代码


                    //string iyear = "20" + Regex.Match(LstData[i], "(?<=\"年\":).*(?=,\"月\")").Value;
                    //if (!string.IsNullOrEmpty(iyear.Trim()))
                    //{
                    //    year = Convert.ToInt32(iyear.Substring(0, iyear.Length - 2));
                    //}
                    //string imonth = Regex.Match(LstData[i], "(?<=\"月\":).*(?=,\"日\")").Value;
                    //if (!string.IsNullOrEmpty(imonth.Trim()))
                    //{
                    //    month = Convert.ToInt32(imonth.Substring(0, imonth.Length - 2));
                    //}
                    //string iday = Regex.Match(LstData[i], "(?<=\"日\":).*(?=,\"时\")").Value;
                    //if (!string.IsNullOrEmpty(iday.Trim()))
                    //{
                    //    day = Convert.ToInt32(iday.Substring(0, iday.Length - 2));
                    //}
                    //string ihour = Regex.Match(LstData[i], "(?<=\"时\":).*(?=,\"分\")").Value;
                    //if (!string.IsNullOrEmpty(ihour.Trim()))
                    //{
                    //    hour = Convert.ToInt32(ihour.Substring(0, ihour.Length - 2));
                    //}
                    //string iminute = Regex.Match(LstData[i], "(?<=\"分\":).*(?=,\"秒\")").Value;
                    //if (!string.IsNullOrEmpty(iminute.Trim()))
                    //{
                    //    minute = Convert.ToInt32(iminute.Substring(0, iminute.Length - 2));
                    //}
                    //string isecond = Regex.Match(LstData[i], "(?<=\"秒\":)[^}]*").Value;
                    //if (!string.IsNullOrEmpty(isecond.Trim()))
                    //{
                    //    second = Convert.ToInt32(isecond.Substring(0, isecond.Length - 2));
                    //}
                    //string DTime = (new DateTime(year, month, day, hour, minute, second)).ToString();
                    //JSList.Add(DTime);

                    //string lt = Regex.Match(LstData[i], "(?<=\"纬度\":).*(?=,\"经度\")").Value;
                    //JSList.Add(lt);

                    //string lg = Regex.Match(LstData[i], "(?<=\"经度\":).*(?=,\"车辆速度\")").Value;
                    //JSList.Add(lg);

                    //string Speed = Regex.Match(LstData[i], "(?<=\"车辆速度\":).*(?=,\"海拔\")").Value;
                    //JSList.Add(Speed);

                    //string High = Regex.Match(LstData[i], "(?<=\"海拔\":).*(?=,\"年\")").Value;
                    //JSList.Add(High);

                    //string Temp = Regex.Match(LstData[i], "(?<=温度)[^}]*").Value;
                    //Temp = Regex.Match(Temp, "(?<=\"value\":\").*(?=\")").Value;
                    //JSList.Add(Temp);

                    //string Hdity = Regex.Match(LstData[i], "(?<=湿度)[^}]*").Value;
                    //Hdity = Regex.Match(Hdity, "(?<=\"value\":\").*(?=\")").Value;
                    //JSList.Add(Hdity);

                    //string Carbon = Regex.Match(LstData[i], "(?<=CO2)[^}]*").Value;
                    //Carbon = Regex.Match(Carbon, "(?<=\"value\":).*").Value;
                    //JSList.Add(Carbon);

                    //string Cbon = Regex.Match(LstData[i], "(?<=\"CO\")[^}]*").Value;
                    //Cbon = Regex.Match(Cbon, "(?<=\"value\":).*").Value;
                    //JSList.Add(Cbon);

                    //string Cascophen = Regex.Match(LstData[i], "(?<=甲醛)[^}]*").Value;
                    //Cascophen = Regex.Match(Cascophen, "(?<=\"value\":).*").Value;
                    //JSList.Add(Cascophen);

                    //string PM1 = Regex.Match(LstData[i], "(?<=PM1.0)[^}]*").Value;
                    //PM1 = Regex.Match(PM1, "(?<=\"value\":).*").Value;
                    //JSList.Add(PM1);

                    //string PM2 = Regex.Match(LstData[i], "(?<=PM2.5)[^}]*").Value;
                    //PM2 = Regex.Match(PM2, "(?<=\"value\":).*").Value;
                    //JSList.Add(PM2);


                    //string PM10 = Regex.Match(LstData[i], "(?<=PM10)[^}]*").Value;
                    //PM10 = Regex.Match(PM10, "(?<=\"value\":).*").Value;
                    //JSList.Add(PM10);

                    //string TVOC = Regex.Match(LstData[i], "(?<=TVOC数据)[^}]*").Value;
                    //TVOC = Regex.Match(TVOC, "(?<=\"value\":).*").Value;
                    //JSList.Add(TVOC);


                    #endregion
                    year = 2000 + (int)Convert.ToSingle(Key_Value["年"]);
                    month = (int)Convert.ToSingle(Key_Value["月"]);
                    day = (int)Convert.ToSingle(Key_Value["日"]);
                    hour = (int)Convert.ToSingle(Key_Value["时"]);
                    minute = (int)Convert.ToSingle(Key_Value["分"]);
                    second = (int)Convert.ToSingle(Key_Value["秒"]);
                    Key_Value.Remove("年");
                    Key_Value.Remove("月");
                    Key_Value.Remove("日");
                    Key_Value.Remove("时");
                    Key_Value.Remove("分");
                    Key_Value.Remove("秒");
                    string DTime = (new DateTime(year, month, day, hour, minute, second)).ToString();



                    foreach (var item in Key_Value)
                    {
                        if (dT.Columns.Contains(item.Key))
                        {
                        }
                        else {
                            dT.Columns.Add(item.Key);
                        }
                    }

                    unitnumber = Key_Value["unitnember"];
                    DataRow dr = dT.NewRow();
                    dr["时间"] = DTime;
                    foreach (var item in Key_Value)
                    {

                        dr[item.Key] = item.Value;
                    }
                    dT.Rows.Add(dr);
                    Form1.form1.progressBar2.Value++;
                    //Console.Write("-");

                }
                #region datatable旧代码


                //dT.Columns.Add("时间");
                //dT.Columns.Add("纬度");
                //dT.Columns.Add("经度");
                //dT.Columns.Add("车辆速度");
                //dT.Columns.Add("海拔");
                //dT.Columns.Add("温度");
                //dT.Columns.Add("湿度");
                //dT.Columns.Add("CO2");
                //dT.Columns.Add("CO");
                //dT.Columns.Add("甲醛");
                //dT.Columns.Add("PM1.0");
                //dT.Columns.Add("PM2.5");
                //dT.Columns.Add("PM10");
                //dT.Columns.Add("TVOC级数");
                //int K = dT.Columns.Count;
                //for (int j = 0; j < (JSList.Count / K); j++)
                //{
                //    DataRow dr = dT.NewRow();
                //    dr["时间"] = JSList[j * K];
                //    dr["纬度"] = JSList[j * K + 1];
                //    dr["经度"] = JSList[j * K + 2];
                //    dr["车辆速度"] = JSList[j * K + 3];
                //    dr["海拔"] = JSList[j * K + 4];
                //    dr["温度"] = JSList[j * K + 5];
                //    dr["湿度"] = JSList[j * K + 6];
                //    dr["CO2"] = JSList[j * K + 7];
                //    dr["CO"] = JSList[j * K + 8];
                //    dr["甲醛"] = JSList[j * K + 9];
                //    dr["PM1.0"] = JSList[j * K + 10];
                //    dr["PM2.5"] = JSList[j * K + 11];
                //    dr["PM10"] = JSList[j * K + 12];
                //    dr["TVOC级数"] = JSList[j * K + 13];
                //    dT.Rows.Add(dr);
                //}
                #endregion
                //string filename = DateTime.Now.ToString("yyyyMMddHHmmss") +"+"+unitnumber+ ".csv";
                string filename = Filename + ".csv";
                string exportPath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase ; //生成CSV路径
                exportPath = exportPath + "Excel导出文件夹" + "\\";
                //var filepath = Server.MapPath("~/Excel导出文件夹/") + filename;
                var filepath = exportPath + filename;


                /////////////////////////////////
                FileInfo fileInfo = new FileInfo(filepath);
                if (!fileInfo.Directory.Exists)
                {
                    fileInfo.Directory.Create();
                }

                FileStream file = new FileStream(filepath, FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(file, System.Text.Encoding.Default);
                string data = "";
                //写出列名称
                for (int i = 0; i < dT.Columns.Count; i++)
                {
                    data += dT.Columns[i].ColumnName.ToString();
                    if (i < dT.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
                //写出各行数据
                Console.WriteLine();
                Form1.form1.label2.Text = "正在写入文件" + Filename ;
                Console.Write(DateTime.Now+"正在写入文件=");
               
                for (int i = 0; i < dT.Rows.Count; i++)
                {
                    data = "";
                    for (int j = 0; j < dT.Columns.Count; j++)
                    {
                        string strs = dT.Rows[i][j].ToString();
                        strs = strs.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                        if (strs.Contains(',') || strs.Contains('"') || strs.Contains('\r') || strs.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                        {
                            strs = string.Format("\"{0}\"", strs);
                        }

                        data += strs;
                        if (j < dT.Columns.Count - 1)
                        {
                            data += ",";
                        }
                    }
                    sw.WriteLine(data);
                    //Form1.form1.progressBar2.Value++;
                    //Console.Write("=");
                }
                sw.Close();
                file.Close();

                //mapStr = "/Excel导出文件夹/" + filename;

                mapStr = filepath;
                //Console.WriteLine();
            }
            Console.WriteLine(DateTime.Now+"文件"+ Filename + "导出成功");
            return mapStr;

        }

        public string chooseResult(string value,string valueTable){
            string Finalresult="";
            if ( valueTable != "无")
            {
                value = Convert.ToString((int)Convert.ToSingle(value));
                string result = valueTable.Replace("||", "|");
                string[] fresult = result.Split('|');
                for (int i = 0; i < fresult.Length; i+=2) {
                    if (value.Equals(fresult[i])) {

                        Finalresult = fresult[i + 1];

                    }

                }
                if (Finalresult.Equals("")) {

                    Finalresult = "错误";
                }
                //faultFlag = fresult[1]; //:配置文件中的第一个为非故障，其他的都为故障
                //int index = result.IndexOf(identity.ToString()[0]);
                //if (index > -1)
                //{
                //    string[] Status = result.Split(identity.ToString()[0]);
                //    string[] St = Status[1].Split('|');
                //    status = St[1];
                //}
                //else
                //{
                //    status = "无效";

                //}

            }
            else
            {
                Finalresult = value;
                //status = Varitem.valueTable;//:故障状态
            }


            return Finalresult;
        }
    }
}
