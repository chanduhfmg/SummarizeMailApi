using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using System.Net;


namespace SummarizeMail.Service
{
    public class ExchangeServiceFactory
    {
        private readonly IConfiguration _config;

        public ExchangeServiceFactory(IConfiguration config)
        {
            _config = config;
        }

        public ExchangeService Create()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            var service = new ExchangeService(ExchangeVersion.Exchange2013_SP1)
            {
                Credentials = new WebCredentials(
                    _config["Exchange:Username"],
                    _config["Exchange:Password"],
                    _config["Exchange:Domain"]
                ),
                Url = new Uri(_config["Exchange:EwsUrl"])
            };

            return service;
        }
    }
}
