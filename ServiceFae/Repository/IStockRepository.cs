using ServiceFae.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServiceFae.Repository
{
    public interface IStockRepository
    {
        IEnumerable<Tb_Wb_StockBE> GetAll();
    }
}