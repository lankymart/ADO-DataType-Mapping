## Data Type Mapping in ADODB ##

The following table shows the ADO Data Type mapping between Visual Basic, Access, SQL Server, Oracle, and the .NET Framework.

| ADO <br /> DataType <br /> Enum	| ADO <br /> DataType <br /> Enum <br /> Value | Mapped <br /> Data Type | SQL <br /> Server	| Size | Access |	Oracle | Visual Basic 6.0 |
| --- | --: | --- | --- | --: | --- | --- | --- |
| adBigInt | 20 | Int64 <br /> SqlDbType.BigInt [^10] <br /> OleDbType.BigInt [^11] <br /> DBTYPE_I8 [^9] |	bigint [^9] | 8 ||| Variant |
| adBinary | 128 | Byte[] <br /> SqlDbType.VarBinary [^10] <br /> OleDbType.Binary [^11] <br /> DBTYPE_BYTES [^9] | binary <br /> timestamp | 50 <br /> 8 |	Binary <br /> LongBinary | Raw [^7] | Variant |
| adBoolean |	11 | Boolean <br /> SqlDbType.Bit [^10] <br /> OleDbType.Boolean [^11] <br /> DBTYPE_BOOL [^9] | bit | 1 <br /> 2 |	Bit <br /> YesNo || Boolean |
| adBSTR | 8 | String <br /> OleDbType.BSTR [^11] ||||||
| adChapter	| 136 |	(DataReader) |||||| 	 
| adChar | 129 | String <br /> SqlDbType.Char [^10] <br /> OleDbType.Char [^11] <br /> DBTYPE_STR [^09]	| char | x || Char | String |
| adCurrency | 6 | Decimal <br /> SqlDbType.Money [^10] <br /> OleDbType.Currency [^11] <br /> DBTYPE_CY [^9] | money <br /> smallmoney | 8 <br /> 4 |	Currency || Currency |
| adDate | 7 | DateTime <br /> OleDbType.DBDate [^11] || 0 | DateTime [^2]	||	Date |
| adDBDate | 133 | DateTime <br /> OleDbType.DBDate [^11] ||||||
| adDBFileTime | 137 | DBFileTime [^11] ||||||
| adDBTime | 134 | DateTime <br /> OleDbType.DBTime [^11] ||||||
| adDBTimeStamp | 135 | DateTime <br /> SqlDbType.DateTime [^10] <br /> OleDbType.DBTimeStamp [^11] <br /> DBTYPE_DBTIMESTAMP [^9] |	datetime <br /> smalldatetime | 8 <br /> 4 | DateTime [^1] |	Date | Date |
| adDecimal	| 14 | Decimal <br /> OleDbType.Decimal [^11] |||| Decimal [^7]	| Variant [^6] |
| adDouble | 5 | Double <br /> SqlDbType.Float [^10] <br /> OleDbType.Double [^11] <br /> DBTYPE_R8 [^9] |	float	| 8 | Double | Float | Double |
| adEmpty	| 0 |	Empty [^11] ||||||	 	 	 	 	 
| adError |	10 | External-Exception <br /> OleDbType.Error [^11] ||||||
| adFileTime | 64 | DateTime <br /> OleDbType.Filetime [^11] ||||||
| adGUID | 72 | Guid <br /> SqlDbType.UniqueIdentifier [^10] <br /> OleDbType.Guid [^11] <br /> DBTYPE_R8 [^9] |	uniqueidentifier [^5] | 16 | Guid <br /> ReplicationID [^2], [^3] | Variant |
| adIDispatch	| 9 | Object <br /> OleDbType.IDispatch [^11] ||||||
| adInteger	| 3 | Int32 <br /> SqlDbType.Int [^10] <br /> OleDbType.Integer [^11] <br /> DBTYPE_I4 [^9] | identity [^4] <br /> int | 4 <br /> 4 | Counter <br /> AutoNumber <br /> LongInteger | Int [^7] | Long |
| adIUnknown | 13 | Object <br /> OleDbType.IUnknown [^11] ||||||
| adLongVarBinary | 205 | Byte[] <br /> SqlDbType.VarBinary [^10] <br /> OleDbType.LongVarBinary [^11] <br /> DBTYPE_BYTES [^9] | image | 2147483647 | OLEObject | Long Raw [^7] <br /> Blob [^8] | Variant |
| adLongVarChar | 201 | String <br /> SqlDbType.VarChar [^10] <br /> OleDbType.LongVarChar [^11] <br /> DBTYPE_STR[^9] | text | 2147483647 | Memo [^1], [^2] <br /> Hyperlink [^1], [^2] | Long [^7] <br /> Clob [^8] |	String |
| adLongVarWChar | 203 | String <br /> SqlDbType.NText [^10] <br /> OleDbType.VarWChar [^11] <br /> DBTYPE_WSTR [^9] | ntext [^5] |	1073741823 | Memo [^3] <br /> Hyperlink [^3] | NClob [^8] |	String |
| adNumeric | 131 | Decimal <br /> SqlDbType.Decimal [^10] <br /> OleDbType.Decimal [^11] <br /> DBTYPE_NUMERIC [^9] | decimal <br /> numeric | 9 | Decimal [^3]	Decimal <br /> Integer <br /> Number <br /> SmallInt | Variant [^6] |
| adPropVariant | 138 | Object <br /> OleDbType.PropVariant [^11] ||||||
| adSingle | 4 | Single <br /> SqlDbType.Real [^10] <br /> OleDbType.Single [^11] <br /> DBTYPE_R4 [^9] |	real | 4 | Single || Single |
| adSmallInt | 2 | Int16 <br /> SqlDbType.SmallInt [^10] <br /> OleDbType.SmallInt [^11] <br /> DBTYPE_I2 [^9] | smallInt | 2 | Integer | Short | Integer |
| adTinyInt | 16 | Byte <br /> OleDbType.TinyInt [^11] ||||||
| adUnsignedBigInt | 21 | UInt64 <br /> OleDbType.UnsignedBigInt [^11] ||||||
| adUnsignedInt	| 19 | UInt32 <br /> OleDbType.UnsignedInt [^11] ||||||	 	 	 	 	 
| adUnsignedSmallInt | 18 | UInt16 <br /> OleDbType.UnsignedSmallInt [^11] ||||||
| adUnsignedTinyInt | 17 | Byte <br /> SqlDbType.TinyInt [^10] <br /> OleDbType.UnsignedTinyInt [^11] <br /> DBTYPE_UI1 [^9] | tinyInt | 1 | Byte || Byte |
| adUserDefined | 132	||||||| 	 	 	 	 	 
| adVarBinary | 204 | Byte[] <br /> SqlDbType.VarBinary [^10] <br /> OleDbType.VarBinary [^11] <br /> DBTYPE_BYTES [^9] | varbinary	| 50 | ReplicationID [^1] | Variant |
| adVarChar | 200 | String <br /> SqlDbType.VarChar [^10] <br /> OleDbType.VarChar [^11] <br /> DBTYPE_STR [^9] | varchar | x | Text [^1], [^2] <br /> LongText	| VarChar | String |
| adVariant | 12 | Object <br /> SqlDbType.Variant [^10] <br /> OleDbType.Variant [^11] <br /> DBTYPE_VARIANT [^9] <br /> DBTYPE_SQLVARIANT [^9] | sql_variant [^9] | 8016 || VarChar2	| Variant |
| adVarNumeric | 139 | OleDbType.VarNumeric [^11] ||||||	 	 	 	 
| adVarWChar | 202 | String <br /> SqlDbType.NVarChar [^10] <br /> OleDbType.VarWChar [^11] <br />DBTYPE_WSTR [^9] | nvarchar [^5] | x | Text [^3] | NVarChar2 | String |
| adWChar	| 130 | String <br /> SqlDbType.NChar [^10] <br /> OleDbType.WChar [^11] <br /> DBTYPE_WSTR [^9] | nchar [^5] | x ||| String |

[^1]: ODBC Driver (3.51.171300): Microsoft Access Driver (*.mdb), Access 97 (3.5 format)
[^2]: OLE DB Provider: Microsoft.Jet.OLEDB.3.51, Access 97 (3.5 format)
[^3]: OLE DB Provider: Microsoft.Jet.OLEDB.4.0 , Access 2000 (4.0 format)
[^4]: OLE DB Provider: SQLOLEDB, SQL Server 6.5
[^5]: OLE DB Provider: SQLOLEDB, SQL Server 7.0 +
[^6]: The VB Decimal data type can only be used within a Variant, that is, you cannot declare a VB variable to be of type Decimal.
[^7]: Oracle 8.0.x Note: DECIMAL and INT are synonyms for NUMBER and NUMBER(10) respectively.
[^8]: Oracle 8.1.x
[^9]: OLE DB Provider: SQLOLEDB, SQL Server 2000 +
[^10]: SQL Server .NET Data Provider (via System.Data.SqlTypes)
[^11]: OLE DB .NET Data Provider (via System.Data.OleDb)

>Note: "User Defined" data types (e.g. ID, TID, EmpID, SysName) are not shown on this diagram.

Attribution to [Carl Prothman](https://web.archive.org/web/20180916001055/http://www.carlprothman.net/ContactMe/tabid/79/Default.aspx) (carl @ prothman.org) - [Original Url](https://web.archive.org/web/20180901201211/http://carlprothman.net/Technology/DataTypeMapping/tabid/97/Default.aspx)
