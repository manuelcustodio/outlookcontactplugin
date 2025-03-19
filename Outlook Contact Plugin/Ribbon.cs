using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Outlook_Contact_Plugin.Ribbon;
using Office = Microsoft.Office.Core;
//using OutlookApp = Microsoft.Office.Interop.Outlook.Application;
//using SysException = System.Exception;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Outlook_Contact_Plugin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        // Clase para representar un contacto (se serializará a JSON)
        public class Contacto
        {
            [JsonProperty("nombre")]
            public string Nombre { get; set; }

            [JsonProperty("apellido")]
            public string Apellido { get; set; }

            [JsonProperty("telefono")]
            public string Telefono { get; set; }
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Outlook_Contact_Plugin.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        // Manejador para el botón de la cinta
        public async void OnSendContactsClicked(Office.IRibbonControl control)
        {
            try
            {
                List<Contacto> contactos = ObtenerContactosSeleccionados();

                if (contactos.Count == 0)
                {
                    MessageBox.Show("No hay contactos seleccionados.");
                    return;
                }

                foreach (Contacto contacto in contactos)
                {
                    await EnviarContactos(contacto);
                }

                MessageBox.Show("Todos los contactos han sido enviados.");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        // Método para obtener los contactos seleccionados en Outlook
        private List<Contacto> ObtenerContactosSeleccionados()
        {
            List<Contacto> lista = new List<Contacto>();

            // Crea una instancia de la aplicación Outlook
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Explorer explorer = outlookApp.ActiveExplorer();

            if (explorer.Selection.Count > 0)
            {
                foreach (object item in explorer.Selection)
                {
                    if (item is ContactItem contact)
                    {
                        lista.Add(new Contacto
                        {
                            Nombre = contact.FirstName,
                            Apellido = contact.LastName,
                            Telefono = !string.IsNullOrEmpty(contact.PrimaryTelephoneNumber) ? contact.PrimaryTelephoneNumber : "No disponible"
                        });
                    }
                }
            }
            return lista;
        }

        // Método asíncrono para enviar los contactos al servicio web
        private async Task EnviarContactos(Contacto contacto)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // Serializa el objeto a JSON
                    string json = JsonConvert.SerializeObject(contacto);
                    HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");

                    // Envía la solicitud POST
                    HttpResponseMessage response = await client.PostAsync("https://www.raydelto.org/agenda.php", content);

                    if (!response.IsSuccessStatusCode)
                    {
                        MessageBox.Show($"Error al enviar el contacto {contacto.Nombre}: {response.StatusCode}");
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error enviando contacto " + contacto.Nombre + ": " + ex.Message);
                }
            }
        }

        // Método para enviar un único contacto
        private async void EnviarContacto(Contacto contacto)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // Serializa el objeto individual, no una lista
                    string json = JsonConvert.SerializeObject(contacto);
                    // Línea temporal para ver el formato del JSON
                    MessageBox.Show(json);

                    HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.PostAsync("https://www.raydelto.org/agenda.php", content);

                    if (response.IsSuccessStatusCode)
                    {
                        MessageBox.Show("Contacto enviado correctamente.");
                    }
                    else
                    {
                        MessageBox.Show($"Error en la petición: {response.StatusCode}");
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("Error enviando contacto: " + ex.Message);
                }
            }
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
