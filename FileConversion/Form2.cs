using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Common.BardCodeHook;

namespace FileConversion
{
    public partial class Form2 : Form
    {
        BardCodeHook bardCodeHook = new BardCodeHook();
        public Form2()
        {
            InitializeComponent();
            this.KeyPreview = true;
            bardCodeHook.Start();
            bardCodeHook.BarCodeEvent += new BardCodeHook.BardCodeDeletegate(bardCodeHook_BarCodeEvent);
        }
        //定义变量
        const int AnimationCount = 80;
        private Point endPosition;
        private int count;
        private void button1_Click(object sender, EventArgs e)
        {
            MouseAction.NativeRECT rect;
            //获取主窗体句柄
            IntPtr ptrTaskbar = MouseAction.FindWindow(null, "Form2");
            if (ptrTaskbar == IntPtr.Zero)
            {
                MessageBox.Show("No windows found!");
                return;
            }
            //获取窗体中"button1"按钮
            IntPtr ptrStartBtn = MouseAction.FindWindowEx(ptrTaskbar, IntPtr.Zero, null, "button1");
            if (ptrStartBtn == IntPtr.Zero)
            {
                MessageBox.Show("No button found!");
                return;
            }
            //获取窗体大小
            MouseAction.GetWindowRect(new HandleRef(this, ptrStartBtn), out rect);
            endPosition.X = (rect.left + rect.right) / 2;
            endPosition.Y = (rect.top + rect.bottom) / 2;
            //判断点击按钮
            if (checkBox1.Checked)
            {
                //选择"查看鼠标运行的轨迹"
                this.count = AnimationCount;
                timer1.Start();
            }
            else
            {
                MouseAction.SetCursorPos(endPosition.X, endPosition.Y);
                MouseAction.mouse_event(MouseAction.MouseEventFlag.LeftDown, 0, 0, 0, UIntPtr.Zero);
                MouseAction.mouse_event(MouseAction.MouseEventFlag.LeftUp, 0, 0, 0, UIntPtr.Zero);
                textBox1.Text = String.Format("{0},{1}", MousePosition.X, MousePosition.Y);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //int stepx = (endPosition.X - MousePosition.X) / count;
            //int stepy = (endPosition.Y - MousePosition.Y) / count;
            //count--;
            //if (count == 0)
            //{
            //    timer1.Stop();
            //    MouseAction.mouse_event(MouseAction.MouseEventFlag.LeftDown, 0, 0, 0, UIntPtr.Zero);
            //    MouseAction.mouse_event(MouseAction.MouseEventFlag.LeftUp, 0, 0, 0, UIntPtr.Zero);
            //}
            //textBox1.Text = String.Format("{0},{1}", MousePosition.X, MousePosition.Y);
            //MouseAction.mouse_event(MouseAction.MouseEventFlag.Move, stepx, stepy, 0, UIntPtr.Zero);

            Point mousePosition = Control.MousePosition;
            //string str = Console.ReadLine();
            //if (!string.IsNullOrEmpty(str))
            //{
            //    label1.Text = str;
            //}

            textBox1.Text = string.Format("X:{0}  Y:{1}", mousePosition.X, mousePosition.Y);

            //MouseAction.mouse_event(MouseAction.MouseEventFlag.LeftDown, mousePosition.X, mousePosition.Y, 0, UIntPtr.Zero);
            //MouseAction.mouse_event(MouseAction.MouseEventFlag.LeftUp, mousePosition.X, mousePosition.Y, 0, UIntPtr.Zero);
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            bardCodeHook.Stop();
            bardCodeHook.BarCodeEvent -= new BardCodeHook.BardCodeDeletegate(bardCodeHook_BarCodeEvent); 
        }
        public void bardCodeHook_BarCodeEvent(BarCodes barCode)
        {
            label1.Text = barCode.KeyName;
        }
    }
}
