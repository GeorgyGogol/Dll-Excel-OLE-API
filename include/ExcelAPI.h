//---------------------------------------------------------------------------

#ifndef EXCELAPI_H
#define EXCELAPI_H

//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <list>
#include <vector.h>
#include <DB.hpp>
#include <ADODB.hpp>
#include <DBGridEh.hpp>
//---------------------------------------------------------------------------
// Copyright (c) 2022 Georgy 'Gogol' Gogolev
//---------------------------------------------------------------------------

extern "C" {

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

	}

	namespace exl {

		enum class ExcelTextAlign : short int {
			xlCenter = -4108, xlBottom = -4107, xlTop = -4160, xlLeft = -4131,
			xlRight = -4152
		};

		enum class FillDirection : char {
			Down = 0, Up, Left, Right
		};

	}

	namespace exl {

		class __declspec(dllexport)TExcelObjectNode {
		public:
			TExcelObjectNode();
			TExcelObjectNode(TExcelObjectNode* pParent);
			TExcelObjectNode(const TExcelObjectNode& src);

		protected:
			~TExcelObjectNode();

		private:
			TExcelObjectNode* Parent;

			std::list<TExcelObjectNode*>Childs;

			void AddChildClass(TExcelObjectNode* child);
			void RemoveChildClass(TExcelObjectNode* child);

		protected:
			TExcelObjectNode* getParentNode()const;

		};

	}

	namespace exl {

		class __declspec(dllexport)TExcelObjectData {
		public:
			TExcelObjectData();
			TExcelObjectData(const Variant& data);
			TExcelObjectData(const TExcelObjectData&);

		protected:
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

			TExcelObject* GetCurrentSelectedChild();

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

	public:

	};

}

extern "C" {

	namespace exl {

		class __declspec(dllexport)
			TExcelCells : public TExcelObjectRangedTemplate<TExcelCells> {
		public:
			TExcelCells(TExcelObject* pParent, const Variant& data);
			TExcelCells(TExcelCells&);
			~TExcelCells();

		private:
			unsigned int dColumn, dRow;

		public:
			unsigned int GetColumnsCount();
			unsigned int GetRowCount();

			TExcelCells* GetCell(unsigned int col, unsigned int row);
			TExcelCells* GetCells(unsigned int startColumn,
				unsigned int startRow, unsigned int endColumn,
				unsigned int endRow);

			TExcelCells* Merge();

			TExcelCells* Insert(const Variant& data);
			TExcelCells* InsertString(const String& data);
			TExcelCells* InsertFormula(String& formula);

			Variant ReadValue();
			String ReadValueString();

			TExcelCells* SetHorizontalAlign(ExcelTextAlign align);
			TExcelCells* SetVerticalAlign(ExcelTextAlign align);

			TExcelCells* SetBorders();

			TExcelCells* SetWidth();
			TExcelCells* SetHeight();
			TExcelCells* AutoSize();

			TExcelCells* SetFormat();
		};

	}

	namespace exl {

		class __declspec(dllexport)
			TExcelTableColumn : public TExcelObjectRangedTemplate<
			TExcelTableColumn>

		{
		public:
			TExcelTableColumn(TExcelObject* pParent, const Variant& data);
			TExcelTableColumn(TExcelTableColumn& src);
			~TExcelTableColumn();

			TExcelTableColumn* SetIdentity(int start, int step);

			TExcelTableColumn* SetHorizontalAlign(ExcelTextAlign align);
			TExcelTableColumn* SetVerticalAlign(ExcelTextAlign align);

			TExcelTableColumn* SetBorders();

			TExcelTableColumn* SetWidth();
			TExcelTableColumn* SetHeight();
			TExcelTableColumn* AutoSize();

			TExcelTableColumn* SetFormat();

		};

		class __declspec(dllexport)
			TExcelTable : public TExcelObjectTemplate<TExcelTable> {
		public:
			TExcelTable(TExcelObject* pSheet, const String& tableName);
			TExcelTable(TExcelObject* pSheet, const Variant& vTable);

			TExcelTable(TExcelObject* pSheet, const Variant& vTable,
				const Variant& vTitle);
			TExcelTable(TExcelObject* pSheet, const Variant& vTable,
				TExcelCells* eTitile);

			~TExcelTable();

		private:

			TExcelCells* Title;

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
			TTableCreator();
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

			typedef std::vector<TExcelTablePivotField>__declspec(dllexport)
				tSettingsArray;

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
		public ITExcelNames {
		public:
			TExcelSheet(TExcelObject* pParent, const Variant& data);
			TExcelSheet(const TExcelSheet&);
			~TExcelSheet();

		protected:

		public:

			TExcelCells* SelectCell(unsigned int col, unsigned int row);

			TExcelCells* SelectCells(unsigned int startColumn,
				unsigned int startRow, unsigned int endColumn,
				unsigned int endRow);

			TExcelCells* SelectColumn(unsigned int column);

			TExcelCells* SelectRow(unsigned int row);

			TExcelCells* InsertDataSet(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				bool needInsertFieldNames = false,
				bool needDisableSet = false);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				const String& tableTitle, const String& tableName,
				bool needDisableSet = false);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				const String& tableTitle, bool needDisableSet = false);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDataSet* dataSet,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDataSet* dataSet,
				const String& tableTitle, const String& tableName,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDataSet* dataSet,
				const String& tableTitle,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDataSet* dataSet,
				bool needDisableSet = false);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDBGridEh* gridEh,
				const String& tableTitle, const String& tableName,
				bool needDisableSet = false);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDBGridEh* gridEh,
				const String& tableTitle, bool needDisableSet = false);

			TExcelTable* CreateTable(unsigned int startColumn,
				unsigned int startRow, TDBGridEh* gridEh,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDBGridEh* gridEh,
				const String& tableTitle, const String& tableName,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDBGridEh* gridEh,
				const String& tableTitle,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDBGridEh* gridEh,
				bool needDisableSet = false);

			TExcelTable* CreatePivotTable(unsigned int startColumn,
				unsigned int startRow, TExcelTable* srcTable,
				TPivotSettings* pivotSettings);

			TExcelTable* GetTable(const String& tableName);

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
			TExcelWorkbook(const TExcelWorkbook&);
			TExcelWorkbook(TExcelObject* pParent, const Variant& data);
			~TExcelWorkbook();

			unsigned int SheetCount();
			TExcelSheet* CreateSheet();
			TExcelSheet* CreateSheet(const String& sheetName);
			TExcelSheet* GetCurrentSheet();
			TExcelSheet* SelectSheet(const String& sheetName);
			TExcelSheet* SelectSheet(unsigned int N);

			TExcelSheet* GetSheet(const String& sheetName);
			TExcelSheet* GetSheet(unsigned int N);

			TExcelTable* CreateTable(TDataSet* dataSet,
				const String& sheetName, const String& tableTitle,
				const String& tableName, bool needDisableSet = false);
			TExcelTable* CreateTable(TDataSet* dataSet,
				const String& sheetName, const String& tableTitle,
				bool needDisableSet = false);
			TExcelTable* CreateTable(TDataSet* dataSet,
				const String& sheetName, bool needDisableSet = false);
			TExcelTable* CreateTable(TDataSet* dataSet,
				bool needDisableSet = false);

			TExcelTable* CreateTable(TDBGridEh* gridEh,
				const String& sheetName, const String& tableTitle,
				const String& tableName, bool needDisableSet = false);
			TExcelTable* CreateTable(TDBGridEh* gridEh,
				const String& sheetName, const String& tableTitle,
				bool needDisableSet = false);
			TExcelTable* CreateTable(TDBGridEh* gridEh,
				const String& sheetName, bool needDisableSet = false);
			TExcelTable* CreateTable(TDBGridEh* gridEh,
				bool needDisableSet = false);

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

			TExcelApp(bool visible);
			TExcelApp(bool visible, unsigned int nSheetsInNewWorkbook);

			TExcelApp(const TExcelApp& src);
			~TExcelApp();

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
