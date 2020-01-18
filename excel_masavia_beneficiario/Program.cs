using System;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using ExcelDataReader;
using Newtonsoft.Json;
using System.Linq;
using System.Collections.Generic;
using Syncfusion.XlsIO;
using System.Xml;
using System.Data;

namespace excel_masavia_beneficiario
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            using (var stream = File.Open(@"/Users/danielhernandez/Downloads/Tarjetas_Personalizadas-2.xlsx", FileMode.Open, FileAccess.Read))
            {
                List<Beneficiary> beneficiaries = new List<Beneficiary>();
                var dataSetConfiguration = new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false
                };
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataTable table = reader.AsDataSet(dataSetConfiguration).Tables["Personalizadas"];
                    reader.Close();
                    foreach (DataRow row in table.Rows)
                    {
                        var rut = EFTEC.RutChile.ConvierteTipoRut(row[2].ToString(), 0, false, false);
                        if (rut != null)
                        {
                            Beneficiary beneficiary = new Beneficiary();
                            beneficiary.Subsidiary = row[0].ToString();
                            beneficiary.SubsidiaryName = row[1].ToString();
                            beneficiary.Rut = int.Parse(rut);
                            beneficiary.Div = EFTEC.RutChile.ObtenerDV(rut);
                            beneficiary.Name = row[3].ToString();
                            beneficiary.SurnameFather = row[4].ToString();
                            beneficiary.SurnameMother = row[5].ToString();
                            beneficiary.Email = row[6].ToString();
                            beneficiary.Error = validateBeneficiary(beneficiary);
                            beneficiaries.Add(beneficiary);
                        }

                    }
                }
            }
        }

        public static String validateBeneficiary(Beneficiary beneficiary)
        {
            String error = "";
            if (String.IsNullOrEmpty(beneficiary.Subsidiary))
            {
                error += "Campo Sucursal Vacío ";
            }
            if (String.IsNullOrEmpty(beneficiary.SubsidiaryName))
            {
                error += "Campo Nombre Sucursal Vacío ";
            }
            if (!EFTEC.RutChile.ValidarRut(beneficiary.Rut.ToString() + beneficiary.Div))
            {
                error += "Campo Rut no es valido ";
            }
            if (String.IsNullOrEmpty(beneficiary.Name))
            {
                error += "Campo Nombre Vacío ";
            }
            else if (beneficiary.Name.Length > 50)
            {
                error += "Campo Nombre Excede el Maximo ";
            }
            if (String.IsNullOrEmpty(beneficiary.SurnameFather))
            {
                error += "Campo Apellido Paterno Vacío ";
            }
            else if(beneficiary.SurnameFather.Length > 50)
            {
                error += "Campo Apellido Paterno Excede el Maximo ";
            }
            if(String.IsNullOrEmpty(beneficiary.SurnameMother))
            {
                error += "Campo Apellido Materno Vacío ";
            }
            else if (beneficiary.SurnameMother.Length > 50)
            {
                error += "Campo Apellido Materno Excede el Maximo ";
            }
            if(String.IsNullOrEmpty(beneficiary.Email))
            {
                error += "Campo Email Vacío ";
            }
            else if(beneficiary.Email.Length > 100)
            {
                error += "Campo Email Excede el maximo ";
            }
            else if (EmailValidation.EmailValidator.Validate(beneficiary.Email))
            {
                error += "Campo Email es invalido ";
            }
            return error;
        }
    }
}
