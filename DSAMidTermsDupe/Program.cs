﻿using System;
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

            public (T item, int priority, string total) Dequeue()
            {
                if (Count == 0)
                {
                    throw new InvalidOperationException("Priority queue is empty :(");
                }

                var item = elements[0];
                elements.RemoveAt(0);
                return item;
            }
        }
        static void Main(string[] args)
        {
            List<string> menuRow = new List<string>();
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

            Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\rbuen\Downloads\Menus.xlsx"); //Make sure this is the correct directory.
            _Worksheet excelSheet = excelBook.Sheets[1];
            Range excelRange = excelSheet.UsedRange;
            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                //List<string> menuRow = new List<string>();
                for (int j = 1; j <= cols; j++)
                {
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    {
                        menuRow.Add(excelRange.Cells[i, j].Value2.ToString());
                    }
                }
                list.Add(menuRow);
            }

            Console.WriteLine("Welcome to Buenaflor's Restaurant!");
            Console.WriteLine("\nHere's the list of our Menu:   ");
            int a = 0;
            for (int i = 0; i < menuRow.Count; i += 3)
            {
                Console.WriteLine("".PadLeft(38, '-'));
                Console.WriteLine($"| {menuRow[i],-5} | {menuRow[i + 1],-15} | {menuRow[i + 2],-5} |");
            }
            Console.WriteLine("".PadLeft(38, '-'));

            int menus = 1;
            int totalPriority = 1;
            PriorityQueue<Stack<string>> orders = new PriorityQueue<Stack<string>>();
            for (int i = 0; i < 3; i++)
            {
                Console.WriteLine($"\n\nWhich line would you take: Customer {i + 1}");
                Console.Write("Type [1] for Priority Lane | [2] for normal lane: ");
                int ans = Convert.ToInt32(Console.ReadLine());
                Console.Write("\nHow many food items are you going to order?: ");
                int quantity = Convert.ToInt32(Console.ReadLine());

                int subtotal = 0;
                Stack<string> currentCustomerOrders = new Stack<string>();
                for (int j = 0; j < quantity; j++)
                {
                    Console.Write("\nChoose the number of the food item you want from the menu: ");
                    int foodnum = Convert.ToInt32(Console.ReadLine());
                    Console.Write("\nInsert Quantity of order: ");
                    int foodquan = Convert.ToInt32(Console.ReadLine());

                    foreach (var menuItem in list)
                    {
                        if (Convert.ToString(foodnum) == (menuItem[0]))
                        {
                            subtotal += foodquan * Convert.ToInt32(menuItem[2]);
                            currentCustomerOrders.Push($"{menuItem[1]} x {foodquan}");
                        }
                    }
                }

                orders.EnqueueSort(currentCustomerOrders, totalPriority, subtotal.ToString());
                Console.WriteLine($"Customer {i + 1} ticket number: {totalPriority:D4}");
                totalPriority += ans == 1 ? 1 : 2;
            }

            Console.WriteLine("\nList of Orders:");
            Console.WriteLine("".PadLeft(38, '-'));

            while (orders.Count > 0)
            {
                var order = orders.Dequeue();
                int totalQuantity = order.item.Count;

                Console.WriteLine("| " + order.priority.ToString("D4") + " |");
                Console.WriteLine("".PadLeft(38, '-'));

                foreach (var food in order.item)
                {
                    Console.WriteLine($"| {food} |");
                    Console.WriteLine("".PadLeft(38, '-'));
                }
                Console.WriteLine($"| Total: {order.total} |");
                Console.WriteLine("".PadLeft(38, '-'));
            }
            Console.ReadKey();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}