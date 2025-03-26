using System.Runtime.Serialization;

namespace Microsoft.Exchange.WebServices.Data.Exceptions;

/// <summary>
/// Represents a TimeZoneDefinition validation error due to missing periods, transitions groups,
/// and transitions. This is common when the server is Zimbra. 
/// </summary>
public class AttributesOnlyTimeZoneException : Exception
{
	public AttributesOnlyTimeZoneException()
		: base()
	{
	}

	public AttributesOnlyTimeZoneException(string message)
		: base(message)
	{
	}

	public AttributesOnlyTimeZoneException(string message, Exception innerException)
		: base(message, innerException)
	{
	}

	protected AttributesOnlyTimeZoneException(SerializationInfo info, StreamingContext context)
		: base(info, context)
	{
	}
}