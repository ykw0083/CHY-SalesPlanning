using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace FT_ADDON.CHY
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            // Price Discount AddOn for EIG
            //PRICE_DISCOUNT.FTPriceDiscount obj = new PRICE_DISCOUNT.FTPriceDiscount();

            // Price List AddOn for EIG
            FT_ADDON.CHY._InitializeEnvironment obj = new FT_ADDON.CHY._InitializeEnvironment();

            // Landed Cost AddOn for GS
            //LANDED_COST.FTLandedCost obj = new FT_ADDON.LANDED_COST.FTLandedCost();

            Application.Run();
        }
    }
}