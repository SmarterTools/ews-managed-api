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

	/// <summary>
	/// Represents an implementation of the IEwsHttpWebResponse interface using HttpResponseMessage.
	/// </summary>
	internal class EwsHttpWebResponse : IEwsHttpWebResponse
    {
        private HttpResponseMessage _response;

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsHttpWebResponse"/> class.
        /// </summary>
        /// <param name="response">The response.</param>
        internal EwsHttpWebResponse(HttpResponseMessage response)
        {
            _response = response;
        }

        #region IEwsHttpWebResponse Members

        /// <summary>
        /// Closes the response stream.
        /// </summary>
        public void Close()
        {
	        Dispose();
        }

        /// <summary>
        /// Gets the stream that is used to read the body of the response from the server.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.IO.Stream"/> containing the body of the response.
        /// </returns>
        public async Task<Stream> GetResponseStream()
        {
            return await _response.Content.ReadAsStreamAsync();
        }

        /// <summary>
        /// Gets the method that is used to encode the body of the response.
        /// </summary>
        /// <returns>A string that describes the method that is used to encode the body of the response.</returns>
        public string ContentEncoding => _response.Content.Headers.ContentEncoding.FirstOrDefault() ?? "";

        /// <summary>
        /// Gets the content type of the response.
        /// </summary>
        /// <returns>A string that contains the content type of the response.</returns>
        public string ContentType => _response.Content.Headers.ContentType?.ToString();

        /// <summary>
        /// Gets the headers that are associated with this response from the server.
        /// </summary>
        /// <returns>A <see cref="T:System.Net.WebHeaderCollection"/> that contains the header information returned with the response.</returns>
        public HttpResponseHeaders Headers => _response.Headers;

        /// <summary>
        /// Gets the URI of the Internet resource that responded to the request.
        /// </summary>
        /// <returns>A <see cref="T:System.Uri"/> that contains the URI of the Internet resource that responded to the request.</returns>
        public Uri ResponseUri => _response.RequestMessage?.RequestUri;

        /// <summary>
        /// Gets the status of the response.
        /// </summary>
        /// <returns>One of the System.Net.HttpStatusCode values.</returns>
        public HttpStatusCode StatusCode => _response.StatusCode;

        /// <summary>
        /// Gets the status description returned with the response.
        /// </summary>
        /// <returns>A string that describes the status of the response.</returns>
        public string StatusDescription => _response.ReasonPhrase;

        /// <summary>
        /// Gets the version of the HTTP protocol that is used in the response.
        /// </summary>
        /// <value></value>
        /// <returns>System.Version that contains the HTTP protocol version of the response.</returns>
        public Version ProtocolVersion => _response.Version;
        #endregion

        #region IDisposable Members

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            _response.Dispose();
        }

        #endregion
    }
}