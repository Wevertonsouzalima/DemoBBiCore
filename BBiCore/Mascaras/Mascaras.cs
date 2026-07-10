// =============================================================================
//  Mascaras.cs  —  Máscaras e validações brasileiras (BBiCore.Mascaras)
// -----------------------------------------------------------------------------
//  Utilitários puros (sem UI) para formatar e validar CPF, CNPJ, telefone, CEP,
//  moeda e data no padrão brasileiro. A regra fica AQUI (reutilizável em model,
//  API, testes); os componentes são apenas a casca visual sobre estes.
//
//  Princípio: EXIBE em português (dd/MM/yyyy, vírgula decimal, máscara), mas o
//  valor de trabalho é o NATIVO/limpo — datas como DateTime (persistidas
//  yyyy-MM-dd), moeda como decimal (ponto), documentos como dígitos. Assim não
//  há conflito ao salvar no banco.
//
//  >>> NOTA PARA REORGANIZAÇÃO: cada tipo em sua #region, nomeada pelo destino.
// =============================================================================

using System.Globalization;
using System.Text;

namespace BBiCore.Mascaras
{
    // #region Utilitário base  ->  destino sugerido: Helpers/

    /// <summary>Funções básicas compartilhadas pelas máscaras.</summary>
    public static class TextoMascara
    {
        /// <summary>Cultura brasileira usada em toda a formatação.</summary>
        public static readonly CultureInfo CulturaBr = CultureInfo.GetCultureInfo("pt-BR");

        /// <summary>Remove tudo o que não for dígito.</summary>
        /// <param name="texto">Texto de entrada.</param>
        /// <returns>Apenas os dígitos.</returns>
        public static string SomenteDigitos(string? texto)
        {
            if (string.IsNullOrEmpty(texto))
                return string.Empty;

            StringBuilder sb = new(texto.Length);

            foreach (char c in texto)
                if (char.IsDigit(c))
                    sb.Append(c);

            return sb.ToString();
        }
    }

    // #endregion

    // #region Documentos: CPF  ->  destino sugerido: Validacao/

    /// <summary>Formatação e validação de CPF.</summary>
    public static class Cpf
    {
        /// <summary>Aplica a máscara 000.000.000-00 sobre os dígitos disponíveis (parcial enquanto digita).</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>CPF formatado até onde há dígitos.</returns>
        public static string Formatar(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);
            if (d.Length > 11)
                d = d[..11];

            StringBuilder sb = new();

            for (int i = 0; i < d.Length; i++)
            {
                if (i == 3 || i == 6)
                    sb.Append('.');
                else if (i == 9)
                    sb.Append('-');

                sb.Append(d[i]);
            }

            return sb.ToString();
        }

        /// <summary>Valida um CPF pelos dígitos verificadores.</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Verdadeiro se válido.</returns>
        public static bool EhValido(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);

            if (d.Length != 11)
                return false;

            bool todosIguais = true;
            for (int i = 1; i < 11; i++)
                if (d[i] != d[0])
                {
                    todosIguais = false;
                    break;
                }

            if (todosIguais)
                return false;

            int dv1 = CalcularDigito(d, 9, 10);
            int dv2 = CalcularDigito(d, 10, 11);

            return dv1 == d[9] - '0' && dv2 == d[10] - '0';
        }

        /// <summary>Calcula um dígito verificador de CPF.</summary>
        /// <param name="digitos">Dígitos do CPF.</param>
        /// <param name="quantidade">Quantos dígitos considerar.</param>
        /// <param name="pesoInicial">Peso do primeiro dígito.</param>
        /// <returns>Dígito verificador (0–9).</returns>
        private static int CalcularDigito(string digitos, int quantidade, int pesoInicial)
        {
            int soma = 0;
            int peso = pesoInicial;

            for (int i = 0; i < quantidade; i++)
            {
                soma += (digitos[i] - '0') * peso;
                peso--;
            }

            int resto = soma % 11;
            return resto < 2 ? 0 : 11 - resto;
        }
    }

    // #endregion

    // #region Documentos: CNPJ  ->  destino sugerido: Validacao/

    /// <summary>Formatação e validação de CNPJ.</summary>
    public static class Cnpj
    {
        /// <summary>Pesos do primeiro dígito verificador.</summary>
        private static readonly int[] PesosDv1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];

        /// <summary>Pesos do segundo dígito verificador.</summary>
        private static readonly int[] PesosDv2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2];

        /// <summary>Aplica a máscara 00.000.000/0000-00 sobre os dígitos (parcial enquanto digita).</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>CNPJ formatado até onde há dígitos.</returns>
        public static string Formatar(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);
            if (d.Length > 14)
                d = d[..14];

            StringBuilder sb = new();

            for (int i = 0; i < d.Length; i++)
            {
                if (i == 2 || i == 5)
                    sb.Append('.');
                else if (i == 8)
                    sb.Append('/');
                else if (i == 12)
                    sb.Append('-');

                sb.Append(d[i]);
            }

            return sb.ToString();
        }

        /// <summary>Valida um CNPJ pelos dígitos verificadores.</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Verdadeiro se válido.</returns>
        public static bool EhValido(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);

            if (d.Length != 14)
                return false;

            bool todosIguais = true;
            for (int i = 1; i < 14; i++)
                if (d[i] != d[0])
                {
                    todosIguais = false;
                    break;
                }

            if (todosIguais)
                return false;

            int dv1 = CalcularDigito(d, PesosDv1);
            int dv2 = CalcularDigito(d, PesosDv2);

            return dv1 == d[12] - '0' && dv2 == d[13] - '0';
        }

        /// <summary>Calcula um dígito verificador de CNPJ conforme os pesos.</summary>
        /// <param name="digitos">Dígitos do CNPJ.</param>
        /// <param name="pesos">Pesos a aplicar.</param>
        /// <returns>Dígito verificador (0–9).</returns>
        private static int CalcularDigito(string digitos, int[] pesos)
        {
            int soma = 0;

            for (int i = 0; i < pesos.Length; i++)
                soma += (digitos[i] - '0') * pesos[i];

            int resto = soma % 11;
            return resto < 2 ? 0 : 11 - resto;
        }
    }

    // #endregion

    // #region Documento genérico (CPF ou CNPJ)  ->  destino sugerido: Validacao/

    /// <summary>Trata um documento que pode ser CPF ou CNPJ conforme a quantidade de dígitos.</summary>
    public static class Documento
    {
        /// <summary>Formata como CPF (até 11 dígitos) ou CNPJ (acima disso).</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Documento formatado.</returns>
        public static string Formatar(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);
            return d.Length <= 11 ? Cpf.Formatar(d) : Cnpj.Formatar(d);
        }

        /// <summary>Valida como CPF (11 dígitos) ou CNPJ (14 dígitos).</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Verdadeiro se válido em qualquer um dos dois formatos.</returns>
        public static bool EhValido(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);

            switch (d.Length)
            {
                case 11:
                    return Cpf.EhValido(d);
                case 14:
                    return Cnpj.EhValido(d);
                default:
                    return false;
            }
        }
    }

    // #endregion

    // #region Telefone e CEP  ->  destino sugerido: Formatacao/

    /// <summary>Formatação de telefone brasileiro (fixo e celular).</summary>
    public static class Telefone
    {
        /// <summary>Aplica (00) 0000-0000 ou (00) 00000-0000 conforme a quantidade de dígitos.</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Telefone formatado.</returns>
        public static string Formatar(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);
            if (d.Length > 11)
                d = d[..11];

            StringBuilder sb = new();
            bool celular = d.Length > 10;

            for (int i = 0; i < d.Length; i++)
            {
                if (i == 0)
                    sb.Append('(');
                else if (i == 2)
                    sb.Append(") ");
                else if ((celular && i == 7) || (!celular && i == 6))
                    sb.Append('-');

                sb.Append(d[i]);
            }

            return sb.ToString();
        }
    }

    /// <summary>Formatação de CEP.</summary>
    public static class Cep
    {
        /// <summary>Aplica 00000-000 sobre os dígitos.</summary>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>CEP formatado.</returns>
        public static string Formatar(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);
            if (d.Length > 8)
                d = d[..8];

            if (d.Length <= 5)
                return d;

            return d[..5] + "-" + d[5..];
        }
    }

    // #endregion

    // #region Moeda e data  ->  destino sugerido: Formatacao/

    /// <summary>Formatação/conversão de moeda no padrão brasileiro.</summary>
    public static class Moeda
    {
        /// <summary>Formata um valor decimal como moeda brasileira.</summary>
        /// <param name="valor">Valor.</param>
        /// <param name="comSimbolo">Se inclui "R$".</param>
        /// <returns>Texto formatado (ex.: "1.234,56" ou "R$ 1.234,56").</returns>
        public static string Formatar(decimal valor, bool comSimbolo = false)
            => valor.ToString(comSimbolo ? "C2" : "N2", TextoMascara.CulturaBr);

        /// <summary>Converte um texto (dígitos como centavos) em decimal. Ex.: "123456" → 1234,56.</summary>
        /// <param name="valor">Texto digitado.</param>
        /// <returns>Valor decimal (0 se vazio).</returns>
        public static decimal DeDigitosCentavos(string? valor)
        {
            string d = TextoMascara.SomenteDigitos(valor);
            if (d.Length == 0)
                return 0m;

            return long.Parse(d, CultureInfo.InvariantCulture) / 100m;
        }
    }

    /// <summary>Formatação/conversão de data no padrão brasileiro.</summary>
    public static class Data
    {
        /// <summary>Formata uma data como dd/MM/yyyy.</summary>
        /// <param name="valor">Data.</param>
        /// <returns>Texto formatado.</returns>
        public static string Formatar(DateTime valor)
            => valor.ToString("dd/MM/yyyy", TextoMascara.CulturaBr);

        /// <summary>Tenta converter um texto dd/MM/yyyy em <see cref="DateTime"/>.</summary>
        /// <param name="valor">Texto no formato brasileiro.</param>
        /// <param name="data">Data convertida, quando válida.</param>
        /// <returns>Verdadeiro se a conversão teve sucesso.</returns>
        public static bool TentarConverter(string? valor, out DateTime data)
            => DateTime.TryParseExact(
                valor,
                ["dd/MM/yyyy", "d/M/yyyy"],
                TextoMascara.CulturaBr,
                DateTimeStyles.None,
                out data);
    }

    // #endregion

    // #region Máscara unificada (para componentes string)  ->  destino sugerido: Formatacao/

    /// <summary>Tipos de máscara aplicáveis a campos de texto (documento, telefone, CEP).</summary>
    public enum TipoMascara
    {
        /// <summary>CPF ou CNPJ conforme a quantidade de dígitos.</summary>
        Documento,

        /// <summary>Somente CPF.</summary>
        Cpf,

        /// <summary>Somente CNPJ.</summary>
        Cnpj,

        /// <summary>Telefone fixo ou celular.</summary>
        Telefone,

        /// <summary>CEP.</summary>
        Cep
    }

    /// <summary>Ponto único de formatação/validação por <see cref="TipoMascara"/>.</summary>
    public static class Mascara
    {
        /// <summary>Formata um texto conforme o tipo de máscara.</summary>
        /// <param name="tipo">Tipo de máscara.</param>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Texto formatado.</returns>
        public static string Formatar(TipoMascara tipo, string? valor)
        {
            switch (tipo)
            {
                case TipoMascara.Documento:
                    return Documento.Formatar(valor);
                case TipoMascara.Cpf:
                    return Cpf.Formatar(valor);
                case TipoMascara.Cnpj:
                    return Cnpj.Formatar(valor);
                case TipoMascara.Telefone:
                    return Telefone.Formatar(valor);
                case TipoMascara.Cep:
                    return Cep.Formatar(valor);
                default:
                    return valor ?? string.Empty;
            }
        }

        /// <summary>Valida um texto conforme o tipo. Retorna nulo quando o tipo não tem validação (telefone/CEP).</summary>
        /// <param name="tipo">Tipo de máscara.</param>
        /// <param name="valor">Texto (com ou sem máscara).</param>
        /// <returns>Verdadeiro/falso para documentos; nulo quando não há regra de validação.</returns>
        public static bool? Validar(TipoMascara tipo, string? valor)
        {
            switch (tipo)
            {
                case TipoMascara.Documento:
                    return Documento.EhValido(valor);
                case TipoMascara.Cpf:
                    return Cpf.EhValido(valor);
                case TipoMascara.Cnpj:
                    return Cnpj.EhValido(valor);
                case TipoMascara.Telefone:
                case TipoMascara.Cep:
                default:
                    return null;
            }
        }

        /// <summary>Chave de máscara usada pelo JS do lado do cliente.</summary>
        /// <param name="tipo">Tipo de máscara.</param>
        /// <returns>Identificador textual (ex.: "doc", "tel", "cep").</returns>
        internal static string ChaveJs(TipoMascara tipo)
        {
            switch (tipo)
            {
                case TipoMascara.Documento:
                    return "doc";
                case TipoMascara.Cpf:
                    return "cpf";
                case TipoMascara.Cnpj:
                    return "cnpj";
                case TipoMascara.Telefone:
                    return "tel";
                case TipoMascara.Cep:
                    return "cep";
                default:
                    return "";
            }
        }
    }

    // #endregion
}
