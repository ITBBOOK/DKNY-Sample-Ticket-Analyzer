using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReader
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            DateTime expirationDate = new DateTime(2024, 5, 31);

            // 获取当前时间
            DateTime currentDate = DateTime.Now;

            // 检查是否超过过期日期
            if (currentDate > expirationDate)
            {
                MessageBox.Show("越权使用，程序退出。", "越权提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());

            }
               
        }
    }
}
