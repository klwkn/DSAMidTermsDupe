using System;
using System.Collections.Generic;
using System.Deployment.Internal;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace DSAMidTermsDupe
{
    internal class Program
    {
        public class PriorityQueue<T>
        {
            private List<(T item, int priority, string total)> elements = new List<(T item, int priority, string total)>();

            public int Count
            {
                get { return elements.Count; }
            }

            public void EnqueueSort(T item, int priority, string total)
            {
                elements.Add((item, priority, total));
                elements.Sort((c, d) => c.priority.CompareTo(d.priority));
            }
            public T Dequeue()
            {
                if (Count == 0)
                {
                    throw new InvalidOperationException("Priority queue is empty :(");
                }

                T item = elements[0].item;
                elements.RemoveAt(0);
                return item;
            }
        }
        static void Main(string[] args)
        {
            List<string> list1 = new List<string>();
            List<List<string>> list = new List<List<string>>();
            Stack<string> customer1 = new Stack<string>();
            Stack<string> customer2 = new Stack<string>();
            Stack<string> customer3 = new Stack<string>();
            Application excelApp = new Application();

            if (excelApp == null)
            {
                ConsoleColor foreground = ConsoleColor.Red;
                Console.WriteLine("EXCEL INVALID!");
                Environment.Exit(1);
            }

            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\rbuen\Downloads\Menus.xlsx");
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                for (int j = 1; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        list1.Add(excelRange.Cells[i, j].Value2.ToString());
                    }
                }
            }

            Console.WriteLine("Welcome to Buenaflor's Restaurant!");
            Console.WriteLine("\nHere's the list of our Menu:   ");
            for (int i = 0; i < list1.Count; i += 3)
            {
                Console.WriteLine("".PadLeft(38, '-'));
                Console.WriteLine($"| {list1[i],-5} | {list1[i + 1],-15} | {list1[i + 2],-5} |");
            }
            Console.WriteLine("".PadLeft(38, '-'));

            int menus = 1;
            for (int i = 0; i < 3; i++)
            {
                Console.WriteLine($"\n\nWhich line would you take: Customer {i + 1}");
                Console.Write("Type [1] for Priority Lane | [2] for normal lane: ");
                int ans = Convert.ToInt32(Console.ReadLine());
                Console.Write("\nHow many food items are you going to order?: ");
                int quantity = Convert.ToInt32(Console.ReadLine());
                int subtotal = 0;

                for (int j = 0; j < quantity; j++)
                {
                    Console.Write("\nChoose the number of the food item you want from the menu: ");
                    int foodnum = Convert.ToInt32((Console.ReadLine()));

                    Console.Write("\nInsert Quantity of order: ");
                    int foodquan = Convert.ToInt32(Console.ReadLine());
                    while (menus < list1.Count)
                    {
                        if (Convert.ToString(foodnum) == (list1[menus]))
                        {
                            subtotal += foodquan * Convert.ToInt32(list1[menus+2]);
                            if (i == 0)
                            {
                                customer1.Push(list1[menus+1]);
                            }
                            else if (i == 1)
                            {
                                customer2.Push(list1[menus+1]);
                            }
                            else
                            {
                                customer3.Push(list1[menus+1]);
                            }
                        }
                        menus+= 3;
                    }
                    menus = 1;
                }
                string custnum = (i + 1).ToString().PadLeft(4, '0');
                Console.WriteLine($"Customer {i + 1} ticket number: {custnum}");
            }

            Console.ReadKey();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}
