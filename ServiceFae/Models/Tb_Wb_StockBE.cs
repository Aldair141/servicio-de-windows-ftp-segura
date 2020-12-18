using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceFae.Models
{
    public class Tb_Wb_StockBE
    {
        public string ModuloId { get; set; }
        public string Local { get; set; }
        public string Productid { get; set; }
        public string ArticIdOld { get; set; }
        public string ColorId { get; set; }
        public string TallaId { get; set; }
        public string ColTall { get; set; }
        public string Talla { get; set; }
        public string Stock { get; set; }
        public string Reserva { get; set; }
        public string StockDisponible { get; set; }
        public string PorcReservaWeb { get; set; }
        public string StockDisponibleWeb { get; set; }
    }
}