// Install-Package ExcelDataReader -Version 3.6.0 -ProjectName TinyExcelReader
// Install-Package System.Text.Encoding.CodePages -Version 5.0.0 -ProjectName TinyExcelReader

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

using ExcelDataReader;

[assembly: CLSCompliant(true)]

namespace TinyExcelReader {

  //-------------------------------------------------------------------------------------------------------------------
  //
  /// <summary>
  /// Standard Excel Reader
  /// </summary>
  //
  //-------------------------------------------------------------------------------------------------------------------

  public static class Excel {
    #region Algorithm

    private static IExcelDataReader CreateReader(Stream stream, ExcelReaderConfiguration config) {
      try {
        return ExcelReaderFactory.CreateReader(stream, config);
      }
      catch (Exception e) {
        throw new TinyExcelReaderException(e.Message, e);
      }
    }

    #endregion Algorithm

    #region Create

    static Excel() {
      Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    #endregion Create

    #region Public

    /// <summary>
    /// Reads entire MS Excel file
    /// </summary>
    /// <param name="filePath">File Path</param>
    /// <param name="password">Password, if any</param>
    /// <returns></returns>
    public static IEnumerable<(string pageName, int pageIndex, string[] row)> ReadLines(
      string filePath,
      string password = null) {

      ExcelReaderConfiguration config = new ExcelReaderConfiguration();

      if (!string.IsNullOrWhiteSpace(password))
        config.Password = password;

      using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read)) {
        using (var reader = CreateReader(stream, config)) {
          int page = 0;

          do {
            while (reader.Read()) {
              string[] row = new string[reader.FieldCount];

              for (int i = 0; i < reader.FieldCount; ++i)
                row[i] = reader.IsDBNull(i) ? null :  $"{reader[i]}";

              yield return (reader.Name, page, row);
            }

            page += 1;
          } while (reader.NextResult());
        }
      }
    }

    /// <summary>
    /// Reads entire MS Excel file
    /// </summary>
    /// <param name="filePath">File Path</param>
    /// <param name="password">Password, if any</param>
    /// <returns></returns>
    public static IEnumerable<(string pageName, int pageIndex, IDataRecord row)> ReadRecords(
      string filePath,
      string password = null) {

      ExcelReaderConfiguration config = new ExcelReaderConfiguration();

      if (!string.IsNullOrWhiteSpace(password))
        config.Password = password;

      using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read)) {
        using (var reader = CreateReader(stream, config)) {
          int page = 0;

          do {
            while (reader.Read())
              yield return (reader.Name, page, reader);

            page += 1;
          } while (reader.NextResult());
        }
      }
    }

    #endregion Public
  }

}
