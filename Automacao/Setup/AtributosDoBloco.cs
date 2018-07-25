namespace TodosBlocos
{
    public class AtributosDoBloco
    {
        private double _X;
        private double _Y;
        private string _nomeEfetivoDoBloco;
        private double _angulo;
        private string _distancia;
        private string _handle;
        private string _nomeBloco;

        public double X { get => _X; set => _X = value; }
        public double Y { get => _Y; set => _Y = value; }
        public string NomeEfetivoDoBloco { get => _nomeEfetivoDoBloco; set => _nomeEfetivoDoBloco = value; }
        public double Angulo { get => _angulo; set => _angulo = value; }
        public string Distancia { get => _distancia; set => _distancia = value; }
        public string Handle { get => _handle; set => _handle = value; }
        public string NomeBloco { get => _nomeBloco; set => _nomeBloco = value; }
    }
    public class TextoElevacao
    {
        private string _elevacaoInicial;
        private string _elevacaoFinal;
        private double _posicaoX;
        private double _posicaoY;

        public string ElevacaoInicial { get => _elevacaoInicial; set => _elevacaoInicial = value; }
        public string ElevacaoFinal { get => _elevacaoFinal; set => _elevacaoFinal = value; }
        public double PosicaoX { get => _posicaoX; set => _posicaoX = value; }
        public double PosicaoY { get => _posicaoY; set => _posicaoY = value; }
    }
}
