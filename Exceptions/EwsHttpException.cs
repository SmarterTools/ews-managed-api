using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.Data
{
	public class EwsHttpException : Exception
	{
		public EwsHttpException(Exception e) : base(e.Message)
		{ }

		public EwsHttpException(HttpResponseMessage response) : base(response.ReasonPhrase)
		{
			IsProtocolError = true;
			Response = response;
		}

		public bool IsProtocolError { get; }
		public HttpResponseMessage Response { get; }
	}
}
