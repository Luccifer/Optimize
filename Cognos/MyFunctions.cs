using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;
using System.Runtime.InteropServices;
using System;
using System.Reflection;
using System.Threading;
using System.Diagnostics;

namespace Cognos
{
    [ComVisible(true)]
    public static class MyFunctions
    {
        public static Class1 main = new Class1();

        [ComVisible(true)]
        [ExcelFunction(Description = "Optimized function", IsMacroType = true, IsThreadSafe = true, SuppressOverwriteError = true, IsClusterSafe = true)]
        public static string fGetVal(string Period, string Period_formula, string Interval, string Actuality, string Company, string ConsType, string Group_Persp, string CurrTypeORCode, string Account, string Movement, string Dim1, string Dim2, string Dim3, string Dim4, string clesVerOrAutoJ, string Is_ContVer, string Form, string Counter_Comp, string Counter_Dim, string JournalNo)
        {
            var current = XlCall.Excel(XlCall.xlfActiveCell) as ExcelReference;
            int row = current.RowFirst;
            int column = current.ColumnFirst;

            var workSpace = new Tuple<int, int>(row, column);
            main.cellAddress = workSpace;

            var perakt = Period + Actuality; var bol = Company;
            //var vkod = CurrTypeORCode;
            var konto = Account; var dim1 = Dim1; var dim2 = Dim2; var dim3 = Dim3; var dim4 = Dim4;
            //var btype
            // var etype
            // var ktypkonc 
            // var vernr
            // var motbol
            // var usrbol
            //var motdim
            // var belopp
            // var trbelopp
            // var vtyp
            // var ino
            // var HID
            // var PFD
            Entity ent = new Entity(perakt, bol, konto, dim1, dim2, dim3, dim4);
            main.addTask(workSpace, ent);

            return "Changes";
        }

        [ComVisible(true)]
        [ExcelFunction(Description = "Insertion Function", IsMacroType = true, IsThreadSafe = true, SuppressOverwriteError = true, IsClusterSafe = true)]
        public static string reload()
        {
            main.connectToDB();
            foreach (Tuple<Entity, Tuple<int, int>> task in main.tasks)
            {
                Entity taskToSend = task.Item1;
                string belopp = main.getBeloppFromDB(taskToSend);

                var reference = new ExcelReference(task.Item2.Item1, task.Item2.Item2);

                ExcelAsyncUtil.Initialize();
                ExcelAsyncUtil.Run("writeInto", new object[] { reference, belopp }, delegate { return writeInto(reference, belopp); });

                //Writer.writeInto(reference, belopp);


                //reference.SetValue(belopp);


                /*ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    reference.SetValue(belopp);
                   
                });*/

            }
            main.closeConnection();
            return "inProcess";
        }

        public static void CheckForMainThread()
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA &&
                !Thread.CurrentThread.IsBackground && !Thread.CurrentThread.IsThreadPoolThread && Thread.CurrentThread.IsAlive)
            {
                MethodInfo correctEntryMethod = Assembly.GetEntryAssembly().EntryPoint;
                StackTrace trace = new StackTrace();
                StackFrame[] frames = trace.GetFrames();
                for (int i = frames.Length - 1; i >= 0; i--)
                {
                    MethodBase method = frames[i].GetMethod();
                    if (correctEntryMethod == method)
                    {
                        return;
                    }
                }
            }

            // throw exception, the current thread is not the main thread...
        }

        public static object writeInto(ExcelReference refer, string value)
        {
           // dynamic application = ExcelDnaUtil.Application;
            ExcelAsyncUtil.QueueAsMacro( () => refer.SetValue(value));
            ExcelAsyncUtil.Uninitialize();
            return value;
        } 
    }
}
