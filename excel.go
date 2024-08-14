package excel

import (
	"fmt"
	"io"
	"sync"

	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
)

type Excel struct {
	file        *excelize.File
	password    string
	sheetCount  int
	sheetPrefix string
	headerRow   int

	once                sync.Once
	importers           []*Importer
	activeSheetNames    []string
	asyncScanWorkerNums int
	humanErrorMsg       bool

	// style
	fieldStyleId int
}

func (e *Excel) doAfterCreateFile(rows []interface{}, initData initData) error {
	if len(e.activeSheetNames) != 0 {
		for _, activeSheetName := range e.activeSheetNames {
			activeSheetIndex, err := e.file.GetSheetIndex(activeSheetName)
			if err != nil {
				return err
			}
			if activeSheetIndex == -1 {
				return errors.Errorf("sheet name %s is invalid or doesn't exist", activeSheetName)
			}
		}
	} else {
		e.activeSheetNames = e.file.GetSheetList()
	}

	if len(e.activeSheetNames) == 0 {
		return errors.New("no sheet exist")
	}

	// now just set one sheet active
	sheetIndex, err := e.file.GetSheetIndex(e.activeSheetNames[_defaultSheetIndex])
	if err != nil {
		return err
	}
	e.file.SetActiveSheet(sheetIndex)

	// set style
	if err := e.initStyle(); err != nil {
		return fmt.Errorf("init excel style error:(%+v)", err)
	}

	if len(rows) != 0 && initData != nil {
		err := initData(rows)
		if err != nil {
			return errors.Wrap(err, "initData")
		}
	}

	if err := e.initImporters(); err != nil {
		return fmt.Errorf("init excel importers error:(%+v)", err)
	}

	return nil
}

func (e *Excel) initStyle() (err error) {
	alignment := &excelize.Alignment{
		Horizontal:  "center",
		Vertical:    "center",
		ShrinkToFit: true,
	}

	e.fieldStyleId, err = e.file.NewStyle(&excelize.Style{
		Alignment: alignment,
	})
	if err != nil {
		return
	}

	return
}

func (e *Excel) initImporters() (err error) {
	e.once.Do(func() {
		for _, sheetName := range e.activeSheetNames {
			root := new(Importer)
			root.value = sheetName
			root.colIndexStart = _colIndexStart
			if root.colIndexEnd, err = e.getSheetLastColIndex(sheetName); err != nil {
				return
			}

			// get sheet headers in merge cells format
			mergeCells, err := e.getHeaders(sheetName)
			if err != nil {
				return
			}

			/*
				set async scan worker nums
			*/
			asyncScanWorkerNums := e.asyncScanWorkerNums
			if asyncScanWorkerNums == 0 {
				asyncScanWorkerNums = _defaultAsyncScanExRowsGoroutineNums
			}
			root.asyncScanWorkerNums = e.asyncScanWorkerNums
			root.withHumanErrorMsg = e.humanErrorMsg

			if root.childImporters, err = buildImporterTree(root, mergeCells); err != nil {
				return
			}
			e.importers = append(e.importers, root)
		}
	})

	return err
}

func (e *Excel) getHeaders(sheet string) (headers []excelize.MergeCell, err error) {
	if e.headerRow == 0 {
		headers, err = e.file.GetMergeCells(sheet)
	} else {
		headers, err = e.getHeadersFromRow(sheet)
	}
	fmt.Println(headers, len(headers))

	return
}

func (e *Excel) getHeadersFromRow(sheet string) (headers []excelize.MergeCell, err error) {
	headerRows, err := e.getHeaderRows(sheet)
	if err != nil {
		err = errors.Wrap(err, "e.getHeaderRows")
		return
	}
	if len(headerRows) == 0 {
		return
	}

	for i, row := range headerRows {
		if len(row) == 0 || row[0] == "" {
			continue
		}

		type headerIndex struct {
			header     string
			start, end int
		}
		headerIndices := make([]headerIndex, 0)
		for j, col := range row {
			if col == "" {
				continue
			}

			l := len(headerIndices)
			if l > 0 {
				headerIndices[l-1].end = j - 1
			}
			headerIndices = append(headerIndices, headerIndex{
				header: col,
				start:  j,
			})
		}
		headerIndices[len(headerIndices)-1].end = len(row) - 1

		for _, headerIndex := range headerIndices {
			var header excelize.MergeCell
			header, err = e.getMergeCell(headerIndex.start+1, headerIndex.end+1, i+1, headerIndex.header)
			if err != nil {
				err = errors.Wrap(err, "e.getMergeCell")
				return
			}
			headers = append(headers, header)
		}

	}
	return
}

func (e *Excel) getHeaderRows(sheet string) ([][]string, error) {
	rows, err := e.file.Rows(sheet)
	if err != nil {
		return nil, err
	}
	results := make([][]string, 0, 64)

	headerRow := e.headerRow
	for rows.Next() && headerRow > 0 {
		row, err := rows.Columns()
		if err != nil {
			break
		}
		results = append(results, row)
		headerRow--
	}
	return results, nil
}

func (e *Excel) getMergeCell(startCol, endCol, row int, value string) (mergeCell excelize.MergeCell, err error) {
	var startAxis, endAxis string
	startAxis, err = excelize.CoordinatesToCellName(startCol, row)
	if err != nil {
		err = errors.Wrap(err, "excelize.CoordinatesToCellName")
		return
	}
	endAxis, err = excelize.CoordinatesToCellName(endCol, row)
	if err != nil {
		err = errors.Wrap(err, "excelize.CoordinatesToCellName")
		return
	}

	mergeCell = make([]string, 2)
	mergeCell[0] = fmt.Sprintf("%s:%s", startAxis, endAxis)
	mergeCell[1] = value
	return
}

func newExcel() (e *Excel) {
	e = new(Excel)
	e.sheetCount = 1
	e.sheetPrefix = _defaultSheetPrefix

	return e
}

func NewExcelFromFile(file string, options ...Option) (e *Excel, err error) {
	e = newExcel()
	for _, option := range options {
		option(e)
	}
	if e.file, err = excelize.OpenFile(file, excelize.Options{Password: e.password}); err != nil {
		return nil, fmt.Errorf("open excel file error, file path:(%s), error:(%+v)", file, err)
	}

	if err = e.doAfterCreateFile(nil, nil); err != nil {
		return nil, err
	}
	return
}

func NewExcelFromReader(reader io.Reader, options ...Option) (e *Excel, err error) {
	e = newExcel()
	for _, option := range options {
		option(e)
	}
	if e.file, err = excelize.OpenReader(reader, excelize.Options{Password: e.password}); err != nil {
		return nil, fmt.Errorf("excel file from reader error, error:(%+v)", err)
	}

	if err = e.doAfterCreateFile(nil, nil); err != nil {
		return nil, err
	}
	return
}

func NewExcelFromData(rows []interface{}, options ...Option) (e *Excel, err error) {
	e = newExcel()
	for _, option := range options {
		option(e)
	}

	e.file = excelize.NewFile()
	for i := 1; i <= e.sheetCount; i++ {
		sheetName := fmt.Sprintf("%s%d", e.sheetPrefix, i)
		e.file.NewSheet(sheetName)
	}

	if e.sheetPrefix != _defaultSheetPrefix {
		// delete default sheet
		e.file.DeleteSheet("Sheet1")
	}

	if err = e.doAfterCreateFile(rows, e.initFromData); err != nil {
		return nil, err
	}
	return
}

/*
*
GetRowsWithHeader get all rows include merge cell header rows
*/
func (e *Excel) GetRowsWithHeader() ([][]string, error) {
	var res [][]string
	for idx, sheet := range e.activeSheetNames {
		if idx == 0 {
			rows, err := e.GetSheetRowsWithHeader(sheet)
			if err != nil {
				return nil, err
			}

			res = append(res, rows...)
			continue
		}

		rows, err := e.GetSheetRowsWithoutHeader(sheet)
		if err != nil {
			return nil, err
		}
		res = append(res, rows...)
	}

	return res, nil
}

/*
*
GetSheetRowsWithHeader get all rows in sheet include merge cell header rows
*/
func (e *Excel) GetSheetRowsWithHeader(sheet string) ([][]string, error) {
	return e.file.GetRows(sheet)
}

/*
*
GetRowsWithoutHeader get all rows except merge cell header rows
*/
func (e *Excel) GetRowsWithoutHeader() ([][]string, error) {
	var res [][]string
	for _, sheetName := range e.activeSheetNames {
		rows, err := e.GetSheetRowsWithoutHeader(sheetName)
		if err != nil {
			return nil, err
		}
		res = append(res, rows...)
	}

	return res, nil
}

/*
*
GetSheetRowsWithoutHeader get all rows in sheet except merge cell header rows
*/
func (e *Excel) GetSheetRowsWithoutHeader(sheet string) ([][]string, error) {
	rowBeginIndex := e.importers[_defaultSheetIndex].getRowsBeginIndex()

	var res [][]string
	rows, err := e.file.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	for i := range rows {
		if i < rowBeginIndex || rows[i] == nil {
			continue
		}
		res = append(res, rows[i])
	}

	return res, nil
}

/*
*
GetCols get columns of the excel
*/
func (e *Excel) GetCols() ([][]string, error) {
	return e.file.GetCols(e.activeSheetNames[_defaultSheetIndex])
}

/*
*
GetSheetCols get columns of the sheet in the excel
*/
func (e *Excel) GetSheetCols(sheet string) ([][]string, error) {
	return e.file.GetCols(sheet)
}

/*
*
GetCellValue get cell value
*/
func (e *Excel) GetCellValue(axis string) (string, error) {
	return e.file.GetCellValue(e.activeSheetNames[_defaultSheetIndex], axis)
}

/*
*
GetSheetCellValue get cell value in sheet
*/
func (e *Excel) GetSheetCellValue(sheet, axis string) (string, error) {
	return e.file.GetCellValue(sheet, axis)
}

/*
GetMergeCells get all merge cells
*/
func (e *Excel) GetMergeCells() ([]excelize.MergeCell, error) {
	return e.file.GetMergeCells(e.activeSheetNames[_defaultSheetIndex])
}

/*
GetSheetMergeCells get all merge cells in sheet
*/
func (e *Excel) GetSheetMergeCells(sheet string) ([]excelize.MergeCell, error) {
	return e.file.GetMergeCells(sheet)
}

func (e *Excel) GetCellHyperLink(axis string) (bool, string, error) {
	return e.file.GetCellHyperLink(e.activeSheetNames[_defaultSheetIndex], axis)
}

func (e *Excel) GetSheetCellHyperLink(sheet, axis string) (bool, string, error) {
	return e.file.GetCellHyperLink(sheet, axis)
}

func (e *Excel) SetCellValue(axis string, value interface{}) error {
	return e.file.SetCellValue(e.activeSheetNames[_defaultSheetIndex], axis, value)
}

func (e *Excel) SetSheetCellValue(sheet, axis string, value interface{}) error {
	return e.file.SetCellValue(sheet, axis, value)
}

func (e *Excel) SetCellColor(axis string, color string) error {
	excelStyle := &excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{color},
		},
	}
	style, err := e.file.NewStyle(excelStyle)
	if err != nil {
		return err
	}
	if err = e.file.SetCellStyle(e.activeSheetNames[_defaultSheetIndex], axis, axis, style); err != nil {
		return err
	}
	return nil
}

func (e *Excel) SetSheetCellColor(sheet, axis string, color string) error {
	excelStyle := &excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{color},
		},
	}
	style, err := e.file.NewStyle(excelStyle)
	if err != nil {
		return err
	}
	if err = e.file.SetCellStyle(sheet, axis, axis, style); err != nil {
		return err
	}
	return nil
}

/*
GetSheetList get sheet list
*/
func (e *Excel) GetSheetList() []string {
	return e.file.GetSheetList()
}

/*
SetActiveSheet set active sheet of the excel
*/
func (e *Excel) SetActiveSheet(sheetName string) (err error) {
	var index int
	if index, err = e.file.GetSheetIndex(sheetName); err != nil {
		return
	}
	if index == -1 {
		return fmt.Errorf("sheet name %s is invalid or doesn't exist", sheetName)
	}
	e.file.SetActiveSheet(index)
	return
}

func (e *Excel) GetFile() *excelize.File {
	return e.file
}

func (e *Excel) GetLastColName() (string, error) {
	lastColIndex, err := e.getLastColIndex()
	if err != nil {
		return "", err
	}
	return excelize.ColumnNumberToName(lastColIndex)
}

func (e *Excel) GetNextColName() (string, error) {
	lastColIndex, err := e.getLastColIndex()
	if err != nil {
		return "", err
	}
	return excelize.ColumnNumberToName(lastColIndex + 1)
}

func (e *Excel) getLastColIndex() (int, error) {
	return e.getSheetLastColIndex(e.activeSheetNames[_defaultSheetIndex])
}

func (e *Excel) getSheetLastColIndex(sheet string) (int, error) {
	cols, err := e.file.GetCols(sheet)
	if err != nil {
		return 0, err
	}
	return len(cols), nil
}

/*
*
ScanExRow scan a excel row to structs
Note: resps must be struct pointer types or ScanExRow will return error
*/
func (e *Excel) ScanExRow(row []string, resps ...interface{}) (scanErrColIndex int, err error) {
	importer := e.importers[_defaultSheetIndex]
	return importer.ScanExRow(row, resps...)
}

func (e *Excel) RelativeScanExRow(row []string, resps ...interface{}) (scanErrColIndex int, err error) {
	importer := e.importers[_defaultSheetIndex]
	return importer.RelativeScanExRow(row, resps...)
}

func (e *Excel) AsyncScanExRows(row [][]string, resps ...interface{}) chan *AsyncScanExRes {
	importer := e.importers[_defaultSheetIndex]
	return importer.AsyncScanExRows(row, resps...)
}

func (e *Excel) SheetImporter(sheet string) *Importer {
	idx := -1
	for i, sheetName := range e.activeSheetNames {
		if sheet == sheetName {
			idx = i
			break
		}
	}

	if idx == -1 {
		return nil
	}

	return e.importers[idx]
}

func (e *Excel) SubImporter(path string) *Importer {
	return e.importers[_defaultSheetIndex].SubImporter(path)
}

func (e *Excel) GetColIndexPos() (colIndexStart, colIndexEnd int) {
	return e.importers[_defaultSheetIndex].GetColIndexPos()
}

func (e *Excel) GetRowIndexPos() (rowIndexStart, rowIndexEnd int) {
	return e.importers[_defaultSheetIndex].GetRowIndexPos()
}

func (e *Excel) IsHeaderConsistent(resps ...interface{}) (isConsistent bool, err error) {
	for _, importer := range e.importers {
		isConsistent, err = importer.IsHeaderConsistent(resps...)
		if err != nil || !isConsistent {
			return
		}
	}

	isConsistent = true
	return
}
