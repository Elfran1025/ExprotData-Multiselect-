using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExprotData
{
   

    public class Info
    {
        public string 纬度 { get; set; }
        public string 经度 { get; set; }
        public string 车辆速度 { get; set; }
        public string 年 { get; set; }
        public string 月 { get; set; }
        public string 日 { get; set; }
        public string 时 { get; set; }
        public string 分 { get; set; }
        public string 秒 { get; set; }
    }

    public class Fault
    {
        public List<故障> 故障 { get; set; }
        public List<报警> 报警 { get; set; }
    }

    public class 整车
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }

    public class 驱动
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }
    public class 燃料电池
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }

    public class 极值
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }
    public class 故障
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }
    public class 报警
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }
    public class 发动机
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }


    public class Vehicle
    {
        public List<整车> 整车 { get; set; }
        public List<驱动> 驱动 { get; set; }
        public List<故障> 故障 { get; set; }
        public List<发动机> 发动机 { get; set; }
        public List<燃料电池> 燃料电池 { get; set; }
        public List<极值> 极值 { get; set; }
        public List<报警> 报警 { get; set; }
    }

    public class Voltage
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }

    public class Temperature
    {
        public string description { get; set; }
        public string unit { get; set; }
        public string valueTable { get; set; }
        public string value { get; set; }
    }

    public class RootObject
    {
        public string unitnember { get; set; }
        public Info info { get; set; }
        public Fault fault { get; set; }
        public Vehicle vehicle { get; set; }
        public List<Voltage> voltage { get; set; }
        public List<Temperature> temperature { get; set; }
    }



}
