using System;
using System.Collections.Generic;

namespace Drenagem.Setup
{
    public static class Constantes
    {
        public static readonly List<String> PrefixoDoNomeDosBlocos = new List<String>
            {"CP_1.20_1.20" , "CP_1.40_1.40" , "CP_1.60_1.60","CP_1.80_1.80","CP_2.00_2.00","CP_2.20_2.20",
            "CP_2.40_2.40","CP_2.60_2.60", "CP_3.20_1.80","CP_3.20_3.20", "CP_3.20_2.40",
            "CDS_1.80_1.40", "CDS_2.60_1.60", "CDS_2.20_1.40", "CDS_2.60_1.20", "CDS_2.80_1.80", "CDS_2.20_1.20",
            "ESG_PV"};

    }

    public static class ConstantesTubulacao
    {
        public static readonly List<String> TubulacaoNomeDosBlocos = new List<String>
        {"TUBO PVC DN 150", "TUBO FF DN 150" , "TUBO FF DN 200" , "TUBO FF DN 250", "TUBO FF DN 300","TUBO FF DN 350","TUBO FF DN 400","TUBO FF DN 450","TUBO FF DN 500",
        "TUBO FF DN 600","TUBO FF DN 700", "TUBO FF DN 800","TUBO FF DN 900", "TUBO FF DN 1000", "TUBO FF DN 1200",
        "TUBO CONC ARMADO DN300", "TUBO CONC ARMADO DN400", "TUBO CONC ARMADO DN500", "TUBO CONC ARMADO DN600", "TUBO CONC ARMADO DN700", "TUBO CONC ARMADO DN800",
        "TUBO CONC ARMADO DN900", "TUBO CONC ARMADO DN1000", "TUBO CONC ARMADO DN1100", "TUBO CONC ARMADO DN1200", "TUBO CONC ARMADO DN1300", "TUBO CONC ARMADO DN1500",
        "TUBO CONC ARMADO DN1750", "TUBO CONC ARMADO DN2000",
        "TUBO CONC SIMPLES DN200","TUBO CONC SIMPLES DN300","TUBO CONC SIMPLES DN400","TUBO CONC SIMPLES DN500","TUBO CONC SIMPLES DN600",
        "GALERIA 2000"};
    }

    public static class ConstantesConexoes
    {
        public static readonly List<String> TubulacaoConexoes = new List<String>
        {"CURVA FF 45 DN150" ,  "CURVA FF 45 DN200" ,  "CURVA FF 45 DN250" ,  "CURVA FF 45 DN300" , "CURVA FF 45 DN350" ,  "CURVA FF 45 DN400" ,  "CURVA FF 45 DN450" ,  "CURVA FF 45 DN500" ,
        "CURVA FF 45 DN600" ,  "CURVA FF 45 DN700" ,  "CURVA FF 45 DN800" ,  "CURVA FF 45 DN900" , "CURVA FF 45 DN1000" ,  "CURVA FF 45 DN1200"};
    }
}
