using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cognos
{
    public class Entity
    {
        public string perakt;
        public string bol;
        public string vkod;
        public string konto;
        public string dim1;
        public string dim2;
        public string dim3;
        public string dim4;
        public string btyp;
        public string etyp;
        public string ktypkonc;
        public int vernr;
        public string motbol;
        public string usrbol;
        public string travkd;
        public string motdim;
        public decimal belopp;
        public decimal trbelopp;
        public string vtyp;
        public int ino;
        public int HID;
        public int PFD;

        public Entity(string perakt, string bol, string konto, string dim1 = "", string dim2 = "", string dim3 = "", string dim4 = "", string vkod = "", string btyp = "", string etyp = "", string ktypkonc = "", int vernr = 0, string motbol = "", string usrbol = "", string travkd = "", string motdim = "", decimal belopp = 0, decimal trbelopp = 0, string vtyp = "", int ino = 0, int HID = 0, int PFD = 0)
        {
            this.perakt = perakt;
            this.bol = bol;
            this.vkod = vkod;
            this.konto = konto;
            this.dim1 = dim1;
            this.dim2 = dim2;
            this.dim3 = dim3;
            this.dim4 = dim4;
            this.btyp = btyp;
            this.etyp = etyp;
            this.ktypkonc = ktypkonc;
            this.vernr = vernr;
            this.motbol = motbol;
            this.usrbol = usrbol;
            this.travkd = travkd;
            this.motdim = motdim;
            this.belopp = belopp;
            this.trbelopp = trbelopp;
            this.vtyp = vtyp;
            this.ino = ino;
            this.HID = HID;
            this.PFD = PFD;
        }
    }
}
