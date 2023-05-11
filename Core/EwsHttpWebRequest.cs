/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

using System.Net.Http.Headers;

namespace Microsoft.Exchange.WebServices.Data
{
	using System;
	using System.IO;
	using System.Net;
	using System.Net.Security;
	using System.Security.Cryptography.X509Certificates;

	/// <summary>
	/// Represents an implementation of the IEwsHttpWebRequest interface that uses HttpWebRequest.
	/// </summary>
	internal class EwsHttpWebRequest : IEwsHttpWebRequest
	{
		private readonly HttpClientHandler _clientHandler;
		private readonly HttpClient _httpClient;

		private volatile bool _isDisposed = false;
		private string _method;
		
		internal EwsHttpWebRequest(Uri uri, bool ignoreSslErrors)
			: this(uri)
		{
			if (ignoreSslErrors)
				_clientHandler.ServerCertificateCustomValidationCallback += (message, certificate2, arg3, arg4) => true;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="EwsHttpWebRequest"/> class.
		/// </summary>
		/// <param name="uri">The URI.</param>
		internal EwsHttpWebRequest(Uri uri)
		{
			Method = "GET";
			RequestUri = uri;

			_clientHandler = new HttpClientHandler { AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip };
			_httpClient = new HttpClient(_clientHandler);
		}

		#region IEwsHttpWebRequest Members

		/// <summary>
		/// Aborts this instance.
		/// </summary>
		public void Abort()
		{
			_httpClient.CancelPendingRequests();
		}
		
		/// <summary>
		/// Returns a response from an internet resource.
		/// </summary>
		/// <returns>A <see cref="Microsoft.Exchange.WebServices.Data.IEwsHttpWebResponse"/> that contains the response from the internet resource.</returns>
		public async Task<IEwsHttpWebResponse> GetResponseAsync()
		{
			return await GetResponseAsync(CancellationToken.None);
		}

		/// <summary>
		/// Returns a response from an internet resource.
		/// </summary>
		/// <param name="token">A cancellation token</param>
		/// <returns>A <see cref="Microsoft.Exchange.WebServices.Data.IEwsHttpWebResponse"/> that contains the response from the internet resource.</returns>
		public async Task<IEwsHttpWebResponse> GetResponseAsync(CancellationToken token)
		{
			var message = new HttpRequestMessage(new HttpMethod(Method), RequestUri);
			message.Content = new StringContent(Content);

			if (!string.IsNullOrEmpty(ContentType))
			{
				message.Content.Headers.ContentType = null;
				message.Content.Headers.TryAddWithoutValidation("Content-Type", ContentType);
			}

			if (!string.IsNullOrEmpty(UserAgent))
			{
				message.Headers.UserAgent.Clear();
				message.Headers.UserAgent.TryParseAdd(UserAgent);
			}

			if (!string.IsNullOrEmpty(Accept))
			{
				message.Headers.Accept.Clear();
				message.Headers.Accept.TryParseAdd(Accept);
			}

			HttpResponseMessage response = null;
			try
			{
				response = await _httpClient.SendAsync(message, token);
			}
			catch (Exception e)
			{
				throw new EwsHttpException(e);
			}

			if (!response.IsSuccessStatusCode)
				throw new EwsHttpException(response);
			return new EwsHttpWebResponse(response);
		}

		/// <summary>
		/// Gets or sets the value of the Accept HTTP header.
		/// </summary>
		/// <returns>The value of the Accept HTTP header. The default value is null.</returns>
		public string Accept
		{
			get;
			set;
		}

		/// <summary>
		/// Gets or sets a value that indicates whether the request should follow redirection responses.
		/// </summary>
		/// <returns>
		/// True if the request should automatically follow redirection responses from the Internet resource; otherwise, false.
		/// The default value is true.
		/// </returns>
		public bool AllowAutoRedirect
		{
			get => _clientHandler.AllowAutoRedirect;
			set => _clientHandler.AllowAutoRedirect = value;
		}

		/// <summary>
		/// Gets or sets the client certificates.
		/// </summary>
		/// <value></value>
		/// <returns>The collection of X509 client certificates.</returns>
		public X509CertificateCollection ClientCertificates
		{
			get;
			set;
		}

		/// <summary>
		/// Gets a or sets the content request content
		/// </summary>
		public string Content
		{
			get;
			set;
		}

		/// <summary>
		/// Gets or sets the value of the Content-type HTTP header.
		/// </summary>
		/// <returns>The value of the Content-type HTTP header. The default value is null.</returns>
		public string ContentType
		{
			get;
			set;
		}

		/// <summary>
		/// Gets or sets the cookie container.
		/// </summary>
		/// <value>The cookie container.</value>
		public CookieContainer CookieContainer
		{
			get => _clientHandler.CookieContainer;
			set => _clientHandler.CookieContainer = value;
		}

		/// <summary>
		/// Gets or sets authentication information for the request.
		/// </summary>
		/// <returns>An <see cref="T:System.Net.ICredentials"/> that contains the authentication credentials associated with the request. The default is null.</returns>
		public ICredentials Credentials
		{
			get => _clientHandler.Credentials;
			set => _clientHandler.Credentials = value;
		}

		/// <summary>
		/// Specifies a collection of the name/value pairs that make up the HTTP headers.
		/// </summary>
		/// <returns>A <see cref="T:System.Net.WebHeaderCollection"/> that contains the name/value pairs that make up the headers for the HTTP request.</returns>
		public HttpRequestHeaders Headers => _httpClient.DefaultRequestHeaders;

		/// <summary>
		/// Gets or sets the method for the request.
		/// </summary>
		/// <returns>The request method to use to contact the Internet resource. The default value is GET.</returns>
		/// <exception cref="T:System.ArgumentException">No method is supplied.-or- The method string contains invalid characters.</exception>
		/// <exception cref="T:System.ObjectDisposedException">The object has been disposed.</exception>
		public string Method
		{
			get => _method;
			set
			{
				if (string.IsNullOrEmpty(value))
					throw new ArgumentException("Value cannot be null or empty", nameof(Method));
				if (IsInvalidHttpMethod(value))
					throw new ArgumentException("Value contains invalid characters", nameof(Method));

				CheckIsDisposed();

				_method = value;
			}
		}

		/// <summary>
		/// Gets or sets proxy information for the request.
		/// </summary>
		public IWebProxy Proxy
		{
			get => _clientHandler.Proxy;
			set => _clientHandler.Proxy = value;
		}

		/// <summary>
		/// Gets or sets a value that indicates whether to send an authenticate header with the request.
		/// </summary>
		/// <returns>true to send a WWW-authenticate HTTP header with requests after authentication has taken place; otherwise, false. The default is false.</returns>
		public bool PreAuthenticate
		{
			get => _clientHandler.PreAuthenticate;
			set => _clientHandler.PreAuthenticate = value;
		}

		/// <summary>
		/// Gets the original Uniform Resource Identifier (URI) of the request.
		/// </summary>
		/// <returns>A <see cref="T:System.Uri"/> that contains the URI of the Internet resource passed to the <see cref="M:System.Net.WebRequest.Create(System.String)"/> method.</returns>
		public Uri RequestUri { get; }

		/// <summary>
		/// Gets or sets the time-out value in milliseconds for the <see cref="M:System.Net.HttpWebRequest.GetResponse"/> and <see cref="M:System.Net.HttpWebRequest.GetRequestStream"/> methods.
		/// </summary>
		/// <returns>The number of milliseconds to wait before the request times out. The default is 100,000 milliseconds (100 seconds).</returns>
		public int Timeout
		{
			get => _httpClient.Timeout.Milliseconds;
			set => _httpClient.Timeout = TimeSpan.FromMilliseconds(value);
		}

		/// <summary>
		/// Gets or sets a <see cref="T:System.Boolean"/> value that controls whether default credentials are sent with requests.
		/// </summary>
		/// <returns>true if the default credentials are used; otherwise false. The default value is false.</returns>
		public bool UseDefaultCredentials
		{
			get => _clientHandler.UseDefaultCredentials;
			set => _clientHandler.UseDefaultCredentials = value;
		}

		/// <summary>
		/// Gets or sets the value of the User-agent HTTP header.
		/// </summary>
		/// <returns>The value of the User-agent HTTP header. The default value is null.The value for this property is stored in <see cref="T:System.Net.WebHeaderCollection"/>. If WebHeaderCollection is set, the property value is lost.</returns>
		public string UserAgent
		{
			get;
			set;
		}

		/// <summary>
		/// Gets or sets if the request to the internet resource should contain a Connection HTTP header with the value Keep-alive
		/// </summary>
		public bool KeepAlive
		{
			get => !(_httpClient.DefaultRequestHeaders.ConnectionClose ?? false);
			set => _httpClient.DefaultRequestHeaders.ConnectionClose = !value;
		}

		/// <summary>
		/// Gets or sets the name of the connection group for the request. 
		/// </summary>
		public string ConnectionGroupName
		{
			get;
			set;
		}

		#endregion

		private void CheckIsDisposed()
		{
			if (_isDisposed)
				throw new ObjectDisposedException(nameof(EwsHttpWebRequest));
		}

		private static readonly char[] INVALID_METHOD_CHARS = { '(', ')', '<', '>', '@', ',', ';', ':', '\\', '"', '\'', '/', '[', ']', '?', '=', '{', '}', ' ', '\t', '\r', '\n' };
		private bool IsInvalidHttpMethod(string method)
		{
			for (var i = 0; i < method.Length; i++)
			{
				if (INVALID_METHOD_CHARS.Contains(method[i]))
					return true;
			}

			return false;
		}

		public void Dispose()
		{
			if (!_isDisposed)
			{
				_isDisposed = true;

				_clientHandler?.Dispose();
				_httpClient?.Dispose();
			}
		}
	}
}