using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    class Program
    {
        static double sred(double[] a)
        {
            double sum = new double();

            for(int i = 0; i < a.GetLength(0); i++)
            {
                sum += a[i];

            }

            sum = sum / a.GetLength(0);

            return sum;

        }

        static double det(double[,] A)
        { // N-размерность матрицы, A-собственно матрица
            double sum = 0;
            if (A.GetLength(0) != 2)
                for (int i = 0; i < A.GetLength(0); i++)
                { //Разложение по первой строке
                    sum += Math.Pow((-1), (i + 2)) * A[0,i] * det(minor(0, i, A));
                }
            else
                sum = A[0,0] * A[A.GetLength(0) - 1, A.GetLength(0) - 1] - A[A.GetLength(0) - 1,0] * A[0, A.GetLength(0) - 1];
            
            return sum;
        }

        static double[,] MultiplN(double[,] a, double n)
        {
            double[,] r = new double[a.GetLength(1), a.GetLength(1)];

            for (int i = 0; i < a.GetLength(0); i++)
            {
                for (int j = 0; j < a.GetLength(1); j++)
                {
                    r[i, j] = a[i, j] * n;
                }
            }

            return r;
        }

        static double[,] minor(int z, int x, double[,] A)
        {
            double[,] C = new double[A.GetLength(0) - 1, A.GetLength(0) - 1];
            
            for (int h = 0, i = 0; i < A.GetLength(0) - 1; i++, h++)
            {
                if (i == z) h++;
                for (int k = 0, j = 0; j < A.GetLength(0) - 1; j++, k++)
                {
                    if (k == x) k++;
                    C[i,j] = A[h,k];
                }
            }
            return C;
        }

        static double[,] Invertible(double[,] a)
        {
            double[,] C = new double[a.GetLength(0), a.GetLength(0)];

            for(int i = 0; i < a.GetLength(0); i++)
            {
                for (int j = 0; j < a.GetLength(0); j++)
                {
                    C[i, j] = Math.Pow(-1, i + j) * det(minor(i, j, a));

                }

            }

            double[,] Inv = MultiplN(Transpose(C), 1/det(a));

            return Inv;
        }

        static double[,] Transpose(double[,] a)
        {
            double[,] r = new double[a.GetLength(1), a.GetLength(0)];

            for (int i = 0; i < r.GetLength(0); i++)
            {
                for (int j = 0; j < r.GetLength(1); j++)
                {
                    r[i, j] = a[j,i];
                }
            }

            return r;
        }

        static double[,] MultiplM(double[,] a, double[,] b)
        {
            if (a.GetLength(1) != b.GetLength(0)) throw new Exception("Матрицы нельзя перемножить");
            double[,] r = new double[a.GetLength(0), b.GetLength(1)];
            for (int i = 0; i < a.GetLength(0); i++)
            {
                for (int j = 0; j < b.GetLength(1); j++)
                {
                    for (int k = 0; k < b.GetLength(0); k++)
                    {
                        r[i, j] += a[i, k] * b[k, j];
                    }
                }
            }
            return r;
        }

        static double[] MultiplV(double[,] a, double[] b)
        {
            double[] r = new double[a.GetLength(0)];

            for (int i = 0; i < a.GetLength(0); i++)
            {
                for (int j = 0; j < a.GetLength(1); j++)
                {
                        r[i] += a[i, j] * b[j];

                }
            }

            return r;
        }

        static void Main(string[] args)
        {
            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"C:\Users\Альмир\Documents\Учеба\Компьютерная обработка эксперементальных данных\Laba1\исходные_данные_для_лаб.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            string[,] XS = new string[76, 5]; // массив значений с листа равен по размеру листу
            string[] yS = new string[76];
            double[,] X = new double[76, 5];
            double[] y = new double[76];
            int[] pv = { 4, 6, 7, 9, 10 };
            double[,] Xt = { { 44.1, 9.3 }, { 41.6, 12 }, { 47.9, 15.6 }, { 46.2, 10.6 }, { 47, 12 } };
            double[] yt = { 13, 13.5, 19, 11.3, 19.4 };
            int n = 76;
            int p = 5;

            for (int j = 0; j < n; j++) // по всем строкам
                yS[j] = ObjWorkSheet.Cells[j + 3, 46].Text.ToString();//считываем текст в строку

            for (int i = 0; i < p; i++) //по всем колонкам
                for (int j = 0; j < n; j++) // по всем строкам
                    XS[j, i] = ObjWorkSheet.Cells[j + 3, pv[i] + 41].Text.ToString();//считываем текст в строку*/
            Console.Write("Вектор Y\n\n");

            for (int j = 0; j < 5; j++)
            {
                Console.Write($"{yt[j]}\t");
                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Матрица X\n\n");

            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    Console.Write($"{Xt[i, j]}\t");

                }

                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Матрица X со свободным членом\n\n");

            double[,] Xft = new double[5, 3];

            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    Xft[i, j] = Xt[i, j];

                }

            }

            for (int i = 0; i < 5; i++)
            {
                Xft[i, 2] = 1;
            }

            for (int i = 0; i < 5; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    Console.Write($"{Xft[i, j]}\t");

                }

                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Оценка а\n\n");

            double[] at = MultiplV(MultiplM(Invertible(MultiplM(Transpose(Xft), Xft)), Transpose(Xft)), yt);

            for (int j = 0; j < at.GetLength(0); j++)
            {
                Console.Write($"{at[j]}\t");
                Console.WriteLine();

            }


            Console.Write("Вектор Y\n\n");

            for (int j = 0; j < n; j++)
            {
                Console.Write($"{yS[j]}\t");
                y[j] = Convert.ToDouble(yS[j]);
                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Матрица X\n\n");

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < p; j++)
                {
                    Console.Write($"{XS[i, j]}\t");
                    X[i, j] = Convert.ToDouble(XS[i, j]);

                }

                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Матрица X со свободным членом\n\n");

            double[,] Xf = new double[76,6];

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < p; j++)
                {
                    Xf[i, j] = X[i,j];

                }

            }

            for (int i = 0; i < n; i++)
            {
                Xf[i, 5] = 1;
            }

            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    Console.Write($"{Xf[i, j]}\t");

                }

                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Оценка а\n\n");

            double[] a = MultiplV( MultiplM( Invertible( MultiplM( Transpose(Xf),Xf ) ), Transpose(Xf)), y);

            for (int j = 0; j < a.GetLength(0); j++)
            {
                Console.Write($"{a[j]}\t");
                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Рассчетное значение y\n\n");

            double[] ocY = MultiplV(Xf,a);

            for (int j = 0; j < n; j++)
            {
                Console.Write($"{y[j]}\t{ocY[j]}\t");
                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Среднее значение\n\n");

            Console.Write($"{sred(y)}\t{sred(ocY)}\t");
            Console.WriteLine();

            Console.ReadKey();
            Console.Write("Вектор оценочных отклонений е\n\n");

            double[] e = new double[76];

            for (int j = 0; j < n; j++)
            {
                e[j] = y[j] - ocY[j];

                Console.Write($"{e[j]}\t");
                Console.WriteLine();

            }

            Console.ReadKey();
            Console.Write("Коэффициент детерминации\n\n");

            double s = new double();
            double s1 = new double();
            double av = sred(y);

            for (int i = 0; i < n; i++)
            {
                s += (ocY[i] - av) * (ocY[i] - av);
                s1 += (y[i] - av) * (y[i] - av);

            }

            double deter = 1 - (s / s1);

            Console.Write($"{deter}\t");

        }
    }
}
