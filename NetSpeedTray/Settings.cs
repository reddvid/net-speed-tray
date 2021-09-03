using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NetSpeedTray
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();

            EnumerateDevices();

            this.FormClosing += Settings_FormClosing;
        }

        private void Settings_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        private void EnumerateDevices()
        {
            cbDevices.Items.Clear();

            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();

            if (adapters.Length != 0)
            {
                foreach (NetworkInterface adapter in adapters)
                {
                    cbDevices.Items.Add(adapter.Description);
                }

                cbDevices.SelectedIndex = Properties.Settings.Default.Device;
            }
        }

        private void cbDevices_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Device = cbDevices.SelectedIndex;
            Properties.Settings.Default.Save();
        }
    }
}
