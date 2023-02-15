//---------------------------------------------------------------------------

#ifndef EXCELAPI_H
#define EXCELAPI_H

//---------------------------------------------------------------------------
// ExcelAPI
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
// ver. 0.2.3.1
//---------------------------------------------------------------------------
#include <list>
#include <vector>

#include <Classes.hpp>
#include <DB.hpp>
#include <ADODB.hpp>
#include <DBGridEh.hpp>
//---------------------------------------------------------------------------

namespace exl {


	enum class XlPivotTableSourceType : short int {
		xlConsolidation = 3, xlDatabase = 1, xlExternal = 2,
		xlPivotTable = -4148
	};

	enum class XlConsolidationFunction : short int {
		xlAverage = -4106, xlCount = -4112, xlCountNums = -4113,

		xlMax = -4136, xlMin = -4139, xlProduct = -4149, xlStDev = -4155,
		xlStDevP = -4156, xlSum = -4157, xlUnknown = 1000, xlVar = -4164,
		xlVarP = -4165
	};

	enum class XlPivotFieldOrientation : unsigned short int {
		xlColumnField = 2, xlDataField = 4, xlHidden = 0, xlPageField = 3,
		xlRowField = 1
	};

	enum class ExcelFormats : char {
		General, Text, Number, Integer, DateShort, DateLong
	};

	enum class ExcelTextAlign : short int {
		xlCenter = -4108, xlBottom = -4107, xlTop = -4160, xlLeft = -4131,
		xlRight = -4152
	};

	enum class XlBordersIndex : char {
		xlDiagonalDown = 5, xlDiagonalUp = 6, xlEdgeBottom = 9, xlEdgeLeft = 7,
		xlEdgeRight = 10, xlEdgeTop = 8, xlInsideHorizontal = 12,
		xlInsideVertical = 11
	};

	enum class XlBorderWeight : short int {
		xlHairline = 1, xlMedium = -4138, xlThick = 4, xlThin = 2
	};

	enum class XlLineStyle : short int {
		xlContinuous = 1, xlDash = -4115, xlDashDot = 4, xlDashDotDot = 5,
		xlDot = -4118, xlDouble = -4119, xlLineStyleNone = -4142,
		xlSlantDashDot = 13
	};

	enum class FillDirection : char {
		Down = 0, Up, Left, Right
	};

	const short int xlNone = -4142;

}

namespace exl {

	template<class T>
	class IFormatManager {
	public:

		virtual IFormatManager<T> * GetFormatInterface() = 0;

		virtual T* SetHorizontalAlign(ExcelTextAlign align) = 0;

		virtual T* SetVerticalAlign(ExcelTextAlign align) = 0;

		virtual T* SetTextWrap(bool state) = 0;

		virtual T* SetFormat(ExcelFormats format) = 0;

		virtual T* SetFormat(const String& format) = 0;

	};

}

namespace exl {

	template<class Ttable, class TdataFrom>
	class ICreateTable {
	public:

		virtual Ttable* CreateTable(unsigned int startColumn,
			unsigned int startRow, TdataFrom* dataSet,
			const String& tableTitle, const String& tableName) = 0;
		virtual Ttable* CreateTable(unsigned int startColumn,
			unsigned int startRow, TdataFrom* dataSet,
			const String& tableTitle) = 0;
		virtual Ttable* CreateTable(unsigned int startColumn,
			unsigned int startRow, TdataFrom* dataSet) = 0;

		__inline Ttable* CreateTable(TdataFrom* dataSet,
			const String& tableTitle, const String& tableName) {
			return CreateTable(1, 1, dataSet, tableTitle, tableName);
		}

		__inline Ttable* CreateTable(TdataFrom* dataSet,
			const String& tableTitle) {
			return CreateTable(1, 1, dataSet, tableTitle);
		}

		__inline Ttable* CreateTable(TdataFrom* dataSet) {
			return CreateTable(1, 1, dataSet);
		}

	};

}

namespace exl {

	template<class Ttable>
	class IGetTable {
	public:

		virtual Ttable* GetTable(const String& tableName) = 0;

	};

}

namespace exl {

	template<class T>
	class IBorderManager {
	public:

		virtual IBorderManager<T> * GetBorderInterface() = 0;

		virtual T* SetBorder(XlBordersIndex border, XlLineStyle style,
			XlBorderWeight weight) = 0;

		virtual T* RemoveBorder(XlBordersIndex border) = 0;

	};

}

extern "C" {

	namespace exl {
		class __declspec(dllexport)TExcelObjectNode {

		public:
			TExcelObjectNode();
			TExcelObjectNode(TExcelObjectNode* pParent);
			TExcelObjectNode(const TExcelObjectNode& src);
			~TExcelObjectNode();

			TExcelObjectNode operator = (const TExcelObjectNode & src);

		private:
			TExcelObjectNode* Parent;

			std::list<TExcelObjectNode*>Childs;

			void AddChildClass(TExcelObjectNode* child);
			void RemoveChildClass(TExcelObjectNode* child);

		protected:
			TExcelObjectNode* getParentNode()const;

			void ClearChilds();

		};

	}

	namespace exl {
		class __declspec(dllexport)TExcelObjectData {
		public:
			TExcelObjectData();
			TExcelObjectData(const Variant& data);
			TExcelObjectData(const TExcelObjectData&);
			~TExcelObjectData();

			Variant vData;
			Variant vDataChild;

			void checkDataValide();

			unsigned int getChildCountByType(const String& oType);
			void seekAndSetDataChild(const String& oType, const String& name);
			void seekAndSetDataChild(const String& oType, unsigned int Num);

		public:

			Variant getVariant();

		};

	}

	namespace exl {

		class __declspec(dllexport)TExcelObject : public TExcelObjectNode,
		public TExcelObjectData {
		public:
			TExcelObject();
			TExcelObject(TExcelObject* pParent, const Variant& data);
			TExcelObject(const TExcelObject& src);
			virtual ~TExcelObject();

		private:

		protected:

		public:
			TExcelObject* GetParent()const;
			Variant GetParentVariant();

			TExcelObject* Show();
			TExcelObject* Hide();

			TExcelObject* SetName(const String& newName);
			String GetName();

		};

	}

}

namespace exl {

	template<class T>
	class __declspec(dllexport)TExcelObjectTemplate : public TExcelObject {
	public:
		TExcelObjectTemplate();
		TExcelObjectTemplate(TExcelObject* pParent, const Variant& data);
		TExcelObjectTemplate(const TExcelObject& src);

	protected:
		~TExcelObjectTemplate();

	public:
		T* Show();
		T* Hide();

		T* SetName(const String& newName);

	};

}

namespace exl {
	template<class T>
	class __declspec(dllexport)
		TExcelObjectRangedTemplate : public TExcelObjectTemplate<T> {
	public:
		TExcelObjectRangedTemplate(TExcelObject* pParent, const Variant& data);
		TExcelObjectRangedTemplate(const TExcelObjectRangedTemplate<T>& src);

	protected:
		~TExcelObjectRangedTemplate();

	private:
		int ColStrToInteger(const String& str)const;

	protected:
		AnsiString ColToStrA(unsigned int ACol);
		unsigned int GetColFromStr(const String& str);
		unsigned int GetRowFromStr(const String& str);

		AnsiString GetRangeString(unsigned int startColumn,
			unsigned int startRow, unsigned int endColumn,
			unsigned int endRow);

		AnsiString GetCellString(unsigned int col, unsigned int row);

		void checkColRow(unsigned int& col, unsigned int& row);

		void selectSingle(unsigned int col, unsigned int row);

		void selectRange(unsigned int startColumn, unsigned int startRow,
			unsigned int endColumn, unsigned int endRow);

		void setSingle(unsigned int col, unsigned int row);

		void setRange(unsigned int startColumn, unsigned int startRow,
			unsigned int endColumn, unsigned int endRow);

	public:

	};

}

extern "C" {

	namespace exl {

		class __declspec(dllexport)
			TExcelCells : public TExcelObjectRangedTemplate<TExcelCells>,
		public IFormatManager<TExcelCells>, public IBorderManager<TExcelCells> {
		public:
			TExcelCells(TExcelObject* pParent, const Variant& data);
			TExcelCells(TExcelCells&);
			~TExcelCells();

		private:

		public:
			unsigned int GetColumnsCount();
			unsigned int GetRowCount();

			void Delete();

			TExcelCells* GetCell(unsigned int col, unsigned int row);

			TExcelCells* GetCells(unsigned int startColumn,
				unsigned int startRow, unsigned int endColumn,
				unsigned int endRow);

			TExcelCells* Insert(const Variant& data);
			TExcelCells* InsertString(const String& data);
			TExcelCells* InsertFormula(String& formula);

			Variant ReadValue();
			String ReadValueString();

			TExcelCells* Merge();

			IFormatManager<TExcelCells> * GetFormatInterface();
			TExcelCells* SetHorizontalAlign(ExcelTextAlign align);
			TExcelCells* SetVerticalAlign(ExcelTextAlign align);
			TExcelCells* SetTextWrap(bool state);
			TExcelCells* SetFormat(ExcelFormats format);
			TExcelCells* SetFormat(const String& format);

			IBorderManager<TExcelCells> * GetBorderInterface();

			TExcelCells* SetBorder(XlBordersIndex border, XlLineStyle style,
				XlBorderWeight weight);

			TExcelCells* RemoveBorder(XlBordersIndex border);

		};

	}

	namespace exl {

		class __declspec(dllexport)
			TExcelTableColumn : public TExcelObjectRangedTemplate<
			TExcelTableColumn>, public IFormatManager<TExcelTableColumn> {
		public:
			TExcelTableColumn(TExcelObject* pParent, const Variant& data);
			TExcelTableColumn(TExcelTableColumn& src);
			~TExcelTableColumn();

			IFormatManager<TExcelTableColumn> * GetFormatInterface();
			TExcelTableColumn* SetHorizontalAlign(ExcelTextAlign align);
			TExcelTableColumn* SetVerticalAlign(ExcelTextAlign align);
			TExcelTableColumn* SetTextWrap(bool state);
			TExcelTableColumn* SetFormat(ExcelFormats format);
			TExcelTableColumn* SetFormat(const String& format);

		};

	}

	namespace exl {

		class __declspec(dllexport)
			TExcelTable : public TExcelObjectRangedTemplate<TExcelTable> {
		public:
			TExcelTable(TExcelObject* pSheet, const Variant& vTable);
			TExcelTable(TExcelObject* pSheet, const Variant& vTable,
				TExcelCells* eTitile);
			~TExcelTable();

		private:
			TExcelCells* Title;

			const unsigned int decideOneStepRows(const unsigned int& nColumns)
				const;

		public:
			String GetTitle();
			TExcelTable* SetTitle(const String& title);

			TExcelCells* GetHeader(unsigned int N);
			TExcelCells* GetHeaders();

			TExcelTableColumn* GetColumn(unsigned int N);
			unsigned int ColumnsCount();

			TExcelCells* GetField(unsigned int col, unsigned int row);

			unsigned int RowsCount();

			TExcelTable* AddRow();
			TExcelTable* AddRow(unsigned int pos);
			TExcelTable* AddRows(unsigned int cnt);
			TExcelTable* AddRows(unsigned int from, unsigned int cnt);

			TExcelTable* AddRows(TDataSet* src,
				const Variant& nullValue = Null());

			TExcelTable* DeleteRow(unsigned int row);

		};

	}

	namespace exl {

		struct __declspec(dllexport)TGridHeader {
			String Caption;
			String FieldName;
			bool Visible;
		};

		class __declspec(dllexport)TGridHeaders {
		public:
			TGridHeaders(TDataSet* gridEh);
			TGridHeaders(TDBGridEh* gridEh);

			typedef std::vector<TGridHeader>::iterator iterator;

		private:
			std::vector<TGridHeader>Headers;

			iterator it;
			unsigned int nVisible;

		public:
			void Add(TGridHeader header);
			unsigned int CountVisible()const;
			unsigned int Count()const;
			unsigned int Deep()const;

			iterator Begin();
			iterator End();
			iterator Next();
			iterator NextVisible();
			bool Eof();

			TGridHeader* CurrentHeader();
			TGridHeader* GetHeader(unsigned int N);

			Variant generateVariant();

		};

	}

	namespace exl {
		class __declspec(dllexport)TTableCreator {
		public:

			TTableCreator(TExcelObject* pSheet, TDataSet* dataSet,
				const String& tableTitle, const String& tableName);
			TTableCreator(TExcelObject* pSheet, TDataSet* dataSet,
				const String& tableTitle);
			TTableCreator(TExcelObject* pSheet, TDataSet* dataSet);

			TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh,
				const String& tableTitle, const String& tableName);
			TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh,
				const String& tableTitle);
			TTableCreator(TExcelObject* pSheet, TDBGridEh* gridEh);

			~TTableCreator();

		private:
			TExcelObject* Sheet;
			TGridHeaders* Headers;
			Variant varData;
			unsigned int nRecords;

			void readDataSet(TDataSet* dataSet);

			void init(TDataSet* dataSet);
			void init(TDBGridEh* gridEh);

			void check(unsigned int& col, unsigned int& row);

		public:
			String Title;
			String TableName;
			String TableStyle;

			void ResetData();

			void PrepareNewData(TExcelObject* sheet, TDataSet* dataSet,
				const String& tableTitle = "",
				const String& tableName = "");
			void PrepareNewData(TExcelObject* sheet, TDBGridEh* gridEh,
				const String& tableTitle = "",
				const String& tableName = "");

			bool CanCreate()const;

			TExcelCells* InsertData(unsigned int col, unsigned int row,
				bool needInsertFieldNames = false);
			TExcelTable* CreateTable(unsigned int col, unsigned int row);

		};

	}

	namespace exl {

		enum PivotDataPlace : char {
			Rows = 0, Columns = 1
		};

		struct __declspec(dllexport)TExcelTablePivotField {
			String ColumnName;
			String Caption;

			XlPivotFieldOrientation Type;
			XlConsolidationFunction Function;
		};

		class __declspec(dllexport)TPivotSettings {
		public:
			TPivotSettings();
			TPivotSettings(const String& pivotName);
			~TPivotSettings();

			typedef std::vector<TExcelTablePivotField>tSettingsArray;

		private:
			String PivotName;

		public:
			String NewSheetName;

			String GetPivotName();
			void SetPivotName(const String& name);

			String PivotTitle;
			tSettingsArray Settings;
			PivotDataPlace DataPlace;
			bool ShowRowTotal;
			bool ShowColumnTotal;

			void Clear();
			void Add(const TExcelTablePivotField& setting);
			void AddRow(const String& сolumnName, const String& caption = "");
			void AddColumn(const String& сolumnName,
				const String& caption = "");
			void AddData(const String& сolumnName, const String& caption = "",
				XlConsolidationFunction function =
				XlConsolidationFunction::xlUnknown);

			unsigned int RowCount()const;
			unsigned int ColumnCount()const;
		};

	}

	namespace exl {

		class __declspec(dllexport)TPivotTableCreator {
		public:
			TPivotTableCreator(TExcelTable* source);

			~TPivotTableCreator();

		private:
			TExcelTable* SourceTable;

			TExcelTable* initPivot(TExcelObject* aimSheet, unsigned int col,
				unsigned int row, TPivotSettings* PivotSettings);

		public:

			void ChangeSourceTable(TExcelTable* sourceNew);

			TExcelTable* CreateTable(TExcelObject* aimSheet, unsigned int col,
				unsigned int row, TPivotSettings* PivotSettings);

		};

	}

	namespace exl {

		class ACreateTableController : public ICreateTable<TExcelTable,
		TDataSet>, public ICreateTable<TExcelTable, TDBGridEh> {
		protected:
			ACreateTableController();

			bool NeedDisableDataSet;

		public:
			void setNeedDisableDataSet(bool isNeedDisable);

		};

	}

	namespace exl {
		class __declspec(dllexport)
			TExcelNameItem : public TExcelObjectTemplate<TExcelNameItem> {
		private:

		public:
			TExcelNameItem(TExcelObject* pParent, const Variant& data);
			TExcelNameItem(const TExcelNameItem& src);
			virtual ~TExcelNameItem();

			TExcelNameItem* SetValue(const Variant& val);
			TExcelNameItem* SetValue(const String& val);
			TExcelNameItem* SetRefers(const String& address);

		};

		class ITExcelNames {
		public:
			virtual TExcelNameItem* GetNameItem(const String& itemName) = 0;
			virtual TExcelNameItem* GetNameItem(unsigned int N) = 0;

			virtual TExcelNameItem* AddNamedItem(const String& itemName) = 0;

		};

	}

	namespace exl {

		class __declspec(dllexport)
			TExcelSheet : public TExcelObjectRangedTemplate<TExcelSheet>,
		public ACreateTableController, public ITExcelNames,
		public IGetTable<TExcelTable> {
		public:
			TExcelSheet(TExcelObject* pParent, const Variant& data);
			TExcelSheet(const TExcelSheet&);
			~TExcelSheet();

			TExcelCells* GetCell(unsigned int col, unsigned int row);
			TExcelCells* GetCell(unsigned int startColumn,
				unsigned int startRow, unsigned int endColumn,
				unsigned int endRow);
			TExcelCells* SelectCell(unsigned int col, unsigned int row);
			TExcelCells* SelectCells(unsigned int startColumn,
				unsigned int startRow, unsigned int endColumn,
				unsigned int endRow);
			TExcelCells* SelectColumn(unsigned int column);
			TExcelCells* SelectColumns(unsigned int colStart,
				unsigned int colEnd);
			TExcelCells* SelectRow(unsigned int row);
			TExcelCells* SelectRows(unsigned int rowStart, unsigned int rowEnd);
			TExcelCells* InsertDataSet(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				bool needInsertFieldNames = false);
			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				const String& tableTitle, const String& tableName);
			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				const String& tableTitle);
			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet);
			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDBGridEh* gridEh,
				const String& tableTitle, const String& tableName);
			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDBGridEh* gridEh,
				const String& tableTitle);
			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDBGridEh* gridEh);
			TExcelTable* GetTable(const String& tableName);
			TExcelTable* CreatePivotTable(unsigned int startColumn,
				unsigned int startRow, TExcelTable* srcTable,
				TPivotSettings* pivotSettings);
			TExcelNameItem* GetNameItem(const String& itemName);
			TExcelNameItem* GetNameItem(unsigned int N);
			TExcelNameItem* AddNamedItem(const String& itemName);

		};

	}

	namespace exl {
		class __declspec(dllexport)
			TExcelWorkbook : public TExcelObjectTemplate<TExcelWorkbook>,
		public ITExcelNames {
		public:
			TExcelWorkbook(TExcelObject* pParent, const Variant& data);
			TExcelWorkbook(const TExcelWorkbook&);
			~TExcelWorkbook();

			unsigned int SheetCount();
			TExcelSheet* CreateSheet();
			TExcelSheet* CreateSheet(const String& sheetName);
			TExcelSheet* GetCurrentSheet();
			TExcelSheet* SelectSheet(const String& sheetName);
			TExcelSheet* SelectSheet(unsigned int N);
			TExcelSheet* GetSheet(const String& sheetName);
			TExcelSheet* GetSheet(unsigned int N);

			TExcelWorkbook* Save();
			TExcelWorkbook* Save(const String filePath = "");

			TExcelNameItem* GetNameItem(const String& itemName);
			TExcelNameItem* GetNameItem(unsigned int N);
			TExcelNameItem* AddNamedItem(const String& itemName);
		};

	}

	namespace exl {

		class __declspec(dllexport)
			TExcelApp : public TExcelObjectTemplate<TExcelApp> {
		public:
			TExcelApp();
			TExcelApp(const TExcelApp& src);
			~TExcelApp();

			TExcelApp(bool visible);
			TExcelApp(bool visible, unsigned int nSheetsInNewWorkbook);

		private:
			bool Notifications;

			void Init();

		public:

			TExcelApp* CreateApp(bool visible);
			TExcelApp* CreateApp(bool visible,
				unsigned int nSheetsInNewWorkbook);

			TExcelApp* AttachApp();
			void DeattachApp();

			TExcelApp* TryAttachApp();

			void Close(bool silent = true);
			void Free();

			TExcelApp* SetExcelNotifications(bool stat);
			TExcelApp* SetSheetsInNewWorkbook(unsigned int N);

			unsigned int WorkbookCount();
			TExcelWorkbook* CreateWorkbook();
			TExcelWorkbook* CreateWorkbook(const String& workbookName);
			TExcelWorkbook* GetCurrentWorkbook();
			TExcelWorkbook* OpenWorkbook(const String& path);

		};

	}

}

#endif
