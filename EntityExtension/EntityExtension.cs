using LSEXT;
using LSSERVICEPROVIDERLib;
using Patholab_Common;
using Patholab_DAL_V1;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace EntityExtension
{
    [ComVisible(true)]
    [ProgId("EntityExtension.EntityExtension")]
    public partial class EntityExtension : IEntityExtension
    {
        INautilusServiceProvider _sp;
        OracleConnection connection;
        OracleCommand cmd;
        DataLayer dal;
        INautilusDBConnection ntlCon;
  
        public ExecuteExtension CanExecute(ref IExtensionParameters Parameters)
        {
            try
            {
                _sp = Parameters["SERVICE_PROVIDER"] as INautilusServiceProvider;
                connect();

                var records = Parameters["RECORDS"];

                while (!records.EOF)
                {
                    string a = records.Fields["U_SAMPLE_MSG_ID"].Value.ToString();
                    long id = long.Parse(a);
                    U_SAMPLE_MSG_USER user = dal.FindBy<U_SAMPLE_MSG_USER>(c => c.U_SAMPLE_MSG_ID == id).SingleOrDefault();

                    if (!user.U_STATUS.ToUpper().Equals("H"))
                    {
                        return ExecuteExtension.exInvisible;
                    }

                    records.MoveNext();
                }

                dal.SaveChanges();
                dal.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
                return ExecuteExtension.exInvisible;
            }

            return ExecuteExtension.exEnabled;
        }

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                _sp = Parameters["SERVICE_PROVIDER"] as INautilusServiceProvider;
                connect();
                var records = Parameters["RECORDS"];

                while (!records.EOF)
                {
                    string a = records.Fields["U_SAMPLE_MSG_ID"].Value.ToString();
                    long id = long.Parse(a);
                    U_SAMPLE_MSG_USER user = dal.FindBy<U_SAMPLE_MSG_USER>(c => c.U_SAMPLE_MSG_ID == id).SingleOrDefault();

                    if (user.U_STATUS.ToUpper().Equals("H"))
                    {
                        user.U_STATUS = "N";
                    }

                    records.MoveNext();
                }

                dal.SaveChanges();
                dal.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void connect()
        {
            ntlCon = Utils.GetNtlsCon(_sp);
            Utils.CreateConstring(ntlCon);
            dal = new DataLayer();
            dal.Connect(ntlCon);
        }
    }
}
