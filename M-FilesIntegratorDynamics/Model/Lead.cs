using System;

namespace DynamicsIntegratorPOC.Model
{
    public class Lead
    {
        public string leadid { get; set; }

        public string fullname { get; set; }

        public int bz_tipo_prospect { get; set; }

        public string bz_razaosocial { get; set; }

        public string bz_cnpj { get; set; }

        public string bz_cpf { get; set; }

        public string bz_cpfcnpjrepresentante { get; set; }

        public int? versionnumber { get; set; }

        public DateTime? createdon { get; set; }

    }
}
