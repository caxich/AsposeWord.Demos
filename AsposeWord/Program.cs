using AsposeWord.Demos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeWord
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = WordTableDemo.Export();
            Console.WriteLine("已成功生成word文件：" + Environment.CurrentDirectory + @"\" + filePath);
            Console.WriteLine("请按任意键继续...");
            Console.ReadKey();
        }
    }
}
