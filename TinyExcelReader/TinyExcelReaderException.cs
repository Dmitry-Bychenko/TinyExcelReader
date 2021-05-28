using System;
using System.Runtime.Serialization;

namespace TinyExcelReader {
  
  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Standard Excel Reader Exception
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public class TinyExcelReaderException : Exception {
    #region Create

    /// <summary>
    /// Standard constructor
    /// </summary>
    public TinyExcelReaderException() 
      : base() { }

    /// <summary>
    /// Standard constructor
    /// </summary>
    public TinyExcelReaderException(string message) 
      : base(message) { }

    /// <summary>
    /// Standard constructor
    /// </summary>
    public TinyExcelReaderException(string message, Exception innerException)
      : base(message ?? innerException?.Message ?? "Excel Reader Error", innerException) { }

    /// <summary>
    /// Standard constructor
    /// </summary>
    protected TinyExcelReaderException(SerializationInfo info, StreamingContext context) 
      : base(info, context) { }

    #endregion Create
  }

}
