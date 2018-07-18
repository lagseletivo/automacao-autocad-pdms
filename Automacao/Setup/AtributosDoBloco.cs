namespace Automacao
{
    public class AtributosDoBloco
    {
        private double _X;
        private double _Y;
        private string _nomeEfetivoDoBloco;
        private double _angulo;
        private string _handle;
        private string _nomeBloco;

        public double X { get => _X; set => _X = value; }
        public double Y { get => _Y; set => _Y = value; }
        public string NomeEfetivoDoBloco { get => _nomeEfetivoDoBloco; set => _nomeEfetivoDoBloco = value; }
        public double Angulo { get => _angulo; set => _angulo = value; }
        public string Handle { get => _handle; set => _handle = value; }
        public string nomeBloco { get => _nomeBloco; set => _nomeBloco = value; }
    }
}
