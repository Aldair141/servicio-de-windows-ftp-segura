using Newtonsoft.Json;
using ServiceFae.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ServiceFae.Repository
{
    public class StockRepository : IStockRepository
    {
        public IEnumerable<Tb_Wb_StockBE> GetAll()
        {
            HttpClient Client = new HttpClient();
            HttpResponseMessage response = Client.GetAsync(ConfigurationManager.AppSettings["rutaStock"]).Result;
            response.EnsureSuccessStatusCode();

            string resultado = response.Content.ReadAsStringAsync().Result;
            List<Tb_Wb_StockBE> lista = JsonConvert.DeserializeObject<List<Tb_Wb_StockBE>>(resultado);

            return lista;
        }
    }
}