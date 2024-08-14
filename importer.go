package excel

import (
	"fmt"
	"reflect"
	"strings"
	"sync"

	"github.com/panjf2000/ants"
	"github.com/pkg/errors"
	"github.com/xuri/excelize/v2"
)

const (
	_defaultSheetIndex                   = 0
	_defaultSheetPrefix                  = "Sheet"
	_colIndexStart                       = 1
	_defaultAsyncScanExRowsGoroutineNums = 100
)

type Importer struct {
	// value of current cell node
	value string
	// beginning col index of current node cell
	colIndexStart int
	// end col index of current node cell
	colIndexEnd int
	// beginning row index of current node cell
	rowIndexStart int
	// end row index of current node cell
	rowIndexEnd int
	// path store the path from root to current node
	path []string
	// children nodes of current node
	childImporters []*Importer

	// leaf nodes of current node
	leafNodes []*Importer
	// goroutine nums for async scan rows
	asyncScanWorkerNums int

	// with human error message
	withHumanErrorMsg bool
}

type AsyncScanExRes struct {
	Resps []interface{}
	Err   error
}

/*
*
childrenColBoundaryIndex return start, end index
start - the node's index in mergeCells which is root's first child
end -  the node's index in mergeCells which is root's last child
*/
func (root *Importer) childrenColBoundaryIndex(mergeCells []excelize.MergeCell) (start, end int, err error) {
	var (
		// the beginning col index of current cell
		colIndexStart int
		// the end col index of current cell
		colIndexEnd int
		// current cell's row index
		rowIndex int
	)

	//set to negative to make sure it also work when there is no child.
	end = -1

	for i, cell := range mergeCells {
		startAxis, endAxis := cell.GetStartAxis(), cell.GetEndAxis()
		if colIndexStart, rowIndex, err = excelize.CellNameToCoordinates(startAxis); err != nil {
			return
		}
		if colIndexEnd, _, err = excelize.CellNameToCoordinates(endAxis); err != nil {
			return
		}
		if rowIndex > root.rowIndexEnd && colIndexStart >= root.colIndexStart && colIndexEnd <= root.colIndexEnd {
			if colIndexStart == root.colIndexStart {
				start = i
			}
			if colIndexEnd == root.colIndexEnd {
				end = i
				break
			}
		}
	}
	return
}

/*
*
buildNodeByCell build a node from a merge cell
*/
func buildNodeByCell(mergeCell excelize.MergeCell) (*Importer, error) {
	var (
		// the beginning col index of current cell
		colIndexStart int
		// the end col index of current cell
		colIndexEnd int

		rowIndexStart int
		rowIndexEnd   int

		err error
	)
	node := new(Importer)
	node.value = mergeCell.GetCellValue()
	startAxis, endAxis := mergeCell.GetStartAxis(), mergeCell.GetEndAxis()
	if colIndexStart, rowIndexStart, err = excelize.CellNameToCoordinates(startAxis); err != nil {
		return nil, err
	}
	if colIndexEnd, rowIndexEnd, err = excelize.CellNameToCoordinates(endAxis); err != nil {
		return nil, err
	}
	node.colIndexStart, node.colIndexEnd = colIndexStart, colIndexEnd
	node.rowIndexStart = rowIndexStart
	node.rowIndexEnd = rowIndexEnd

	return node, nil
}

/*
*
buildImporterTree build a excel tree based by mergeCells
*/
func buildImporterTree(root *Importer, mergeCells []excelize.MergeCell) ([]*Importer, error) {
	if root == nil {
		return nil, nil
	}
	var (
		// the node's index in mergeCells which is root's first child
		start int
		// the node's index in mergeCells which is root's last child
		end = -1
		err error
	)
	if start, end, err = root.childrenColBoundaryIndex(mergeCells); err != nil {
		return nil, err
	}
	if end < 0 {
		return nil, nil
	}
	var (
		res  []*Importer
		node *Importer
	)
	for i := start; i <= end; i++ {
		node, err = buildNodeByCell(mergeCells[i])
		if err != nil || node == nil {
			return nil, err
		}

		// children's async scan worker nums just inherit root
		node.asyncScanWorkerNums = root.asyncScanWorkerNums
		node.withHumanErrorMsg = root.withHumanErrorMsg

		// children's path
		node.path = append(node.path, root.path...)
		node.path = append(node.path, mergeCells[i].GetCellValue())

		if node.childImporters, err = buildImporterTree(node, mergeCells); err != nil {
			return nil, err
		}
		res = append(res, node)
	}
	return res, nil
}

/*
*
getRowsBeginIndex return the beginning row index of excel (except of mergeCell headers)
*/
func (root *Importer) getRowsBeginIndex() int {
	if root == nil {
		return 0
	}
	im := root
	for len(im.childImporters) != 0 {
		im = im.childImporters[0]
	}
	return im.rowIndexEnd
}

/*
*
getLeafNodes get the leaf nodes of tree
*/
func (root *Importer) getLeafNodes() []*Importer {
	if root == nil {
		return nil
	}
	var res []*Importer
	if len(root.childImporters) == 0 {
		res = append(res, root)
		return res
	}
	for _, im := range root.childImporters {
		res = append(res, im.getLeafNodes()...)
	}
	return res
}

/*
*
GetColIndexPos return the col start index and col end index of root cell
*/
func (root *Importer) GetColIndexPos() (colIndexStart, colIndexEnd int) {
	return root.colIndexStart, root.colIndexEnd
}

/*
*
GetColIndexPos return the row start index and row end index of root cell
*/
func (root *Importer) GetRowIndexPos() (rowIndexStart, rowIndexEnd int) {
	return root.rowIndexStart, root.rowIndexEnd
}

func (root *Importer) SubImporter(path string) *Importer {
	paths := strings.Split(path, "|")
	if len(root.path) > 0 {
		return root.subImporter(append([]string{root.value}, paths...))
	}
	return root.subImporter(paths)
}

/*
Children get children of current node
*/
func (root *Importer) Children() []*Importer {
	return root.childImporters
}

/*
*
SubImporter return the sub importer by excel path
*/
func (root *Importer) subImporter(path []string) *Importer {
	if len(path) == 0 {
		return nil
	}
	if len(path) == 1 {
		if path[0] == root.value {
			return root
		}
		return nil
	}
	// if root is the root of the excel tree, it's a fake node, it's path is nil, so handle it special,
	// no need to consume the path param, just go on.
	if root.path == nil || root.value == path[0] {
		if root.path != nil {
			path = path[1:]
		}
		for _, node := range root.childImporters {
			if importer := node.subImporter(path); importer != nil {
				return importer
			}
		}
	}
	return nil
}

/*
*
ScanExRow scan a excel row to structs
Note: resps must be struct pointer types or ScanExRow will return error
*/
func (root *Importer) ScanExRow(row []string, resps ...interface{}) (scanErrColIndex int, err error) {
	defer func() {
		if p := recover(); p != nil {
			err = fmt.Errorf("ScanExRow: internal error: %v", p)
		}
	}()
	if len(root.leafNodes) == 0 {
		root.leafNodes = root.getLeafNodes()
	}
	leafNodesLength := len(root.leafNodes)
	rowLength := len(row)
	if leafNodesLength < rowLength {
		row = row[rowLength-leafNodesLength:]
	}

	if rowLength < leafNodesLength {
		for i := 0; i < leafNodesLength-rowLength; i++ {
			row = append(row, "")
		}
	}

	for _, resp := range resps {
		v := reflect.ValueOf(resp).Elem()
		for i := 0; i < reflect.Indirect(v).NumField(); i++ {
			field := reflect.Indirect(v).Type().Field(i)
			tag := field.Tag.Get("ex")
			path := strings.Split(tag, "|")
			for j, leafNode := range root.leafNodes {
				if reflect.DeepEqual(leafNode.path, path) {
					var setValue interface{}
					setValue, err = reflect.Indirect(v).Field(i).Interface().(ImportField).Translate(row[j], leafNode.colIndexStart)
					if err != nil {
						if root.withHumanErrorMsg {
							fmt.Printf("scanExRow scan row:(%+v) error:(%+v)", row, err)
							return leafNode.colIndexStart, errors.Errorf("%s 单元格填写错误，请检查", strings.Join(leafNode.path, "-"))
						}
						return leafNode.colIndexStart, err
					}
					reflect.Indirect(v).Field(i).Set(reflect.ValueOf(setValue))
					break
				}
			}
		}
	}
	return
}

/*
*
RelativeScanExRow scan a excel row to structs, but not like ScanExRow, RelativeScanExRow scan a row by relative ex path
ex: a leaf node's path is `ex:"a|b|c"`, you difine a struct field test which has a tag `ex:"c"` or a tag `ex:"b|c"`
it can scan because your field ex path match the leaf node's behind path
Note: resps must be struct pointer types or ScanExRow will return error
*/
func (root *Importer) RelativeScanExRow(row []string, resps ...interface{}) (scanErrColIndex int, err error) {
	defer func() {
		if p := recover(); p != nil {
			err = fmt.Errorf("ScanExRow: internal error: %v", p)
		}
	}()
	if len(root.leafNodes) == 0 {
		root.leafNodes = root.getLeafNodes()
	}
	leafNodesLength := len(root.leafNodes)
	rowLength := len(row)
	if leafNodesLength < rowLength {
		row = row[rowLength-leafNodesLength:]
	}

	if rowLength < leafNodesLength {
		for i := 0; i < leafNodesLength-rowLength; i++ {
			row = append(row, "")
		}
	}

	for _, resp := range resps {
		v := reflect.ValueOf(resp).Elem()
		for i := 0; i < reflect.Indirect(v).NumField(); i++ {
			field := reflect.Indirect(v).Type().Field(i)
			tag := field.Tag.Get("ex")
			path := strings.Split(tag, "|")
			for j, leafNode := range root.leafNodes {
				if reflect.DeepEqual(leafNode.path[len(leafNode.path)-len(path):], path) {
					var setValue interface{}
					setValue, err = reflect.Indirect(v).Field(i).Interface().(ImportField).Translate(row[j], leafNode.colIndexStart)
					if err != nil {
						if root.withHumanErrorMsg {
							fmt.Printf("scanExRow scan row:(%+v) error:(%+v)", row, err)
							return leafNode.colIndexStart, errors.Errorf("%s 单元格填写错误，请检查", strings.Join(leafNode.path, "-"))
						}
						return leafNode.colIndexStart, err
					}
					reflect.Indirect(v).Field(i).Set(reflect.ValueOf(setValue))
					break
				}
			}
		}
	}
	return
}

/*
*
asyncScanRow scan a row to resps async, put the scaned resps into channel
*/
func (root *Importer) asyncScanRow(row []string, ch chan *AsyncScanExRes, resps ...interface{}) {
	var err error
	asyncScanExRes := new(AsyncScanExRes)
	defer func() {
		if p := recover(); p != nil {
			err = fmt.Errorf("asyncScanRow: internal error: %v", p)
			asyncScanExRes.Err = err
			ch <- asyncScanExRes
		}
	}()
	if len(root.leafNodes) == 0 {
		root.leafNodes = root.getLeafNodes()
	}
	leafNodesLength := len(root.leafNodes)
	rowLength := len(row)
	if leafNodesLength < rowLength {
		row = row[rowLength-leafNodesLength:]
	}

	if rowLength < leafNodesLength {
		for i := 0; i < leafNodesLength-rowLength; i++ {
			row = append(row, "")
		}
	}
	for _, resp := range resps {
		v := reflect.ValueOf(resp).Elem()
		for i := 0; i < reflect.Indirect(v).NumField(); i++ {
			field := reflect.Indirect(v).Type().Field(i)
			tag := field.Tag.Get("ex")
			path := strings.Split(tag, "|")
			for j, leafNode := range root.leafNodes {
				if reflect.DeepEqual(leafNode.path, path) {
					var setValue interface{}
					setValue, err = reflect.Indirect(v).Field(i).Interface().(ImportField).Translate(row[j], leafNode.colIndexStart)
					if err != nil {
						asyncScanExRes.Err = err
						if root.withHumanErrorMsg {
							fmt.Printf("asyncScanRow scan row:(%+v) error:(%+v)", row, err)
							asyncScanExRes.Err = errors.Errorf("%s 单元格填写错误，请检查", strings.Join(leafNode.path, "-"))
						}
						ch <- asyncScanExRes
					} else {
						reflect.Indirect(v).Field(i).Set(reflect.ValueOf(setValue))
					}

					break
				}
			}
		}
		asyncScanExRes.Resps = append(asyncScanExRes.Resps, resp)
	}
	ch <- asyncScanExRes
}

/*
*
AsyncScanExRows scan rows to resps async
*/
func (root *Importer) AsyncScanExRows(rows [][]string, resps ...interface{}) chan *AsyncScanExRes {
	pool, _ := ants.NewPool(root.asyncScanWorkerNums)

	ch := make(chan *AsyncScanExRes, len(rows))
	var wg sync.WaitGroup
	for i := range rows {
		wg.Add(1)
		var respParams []interface{}
		for _, resp := range resps {
			//st := resp
			respParams = append(respParams, reflect.New(reflect.Indirect(reflect.ValueOf(resp).Elem()).Type()).Interface())
		}
		index := i
		_ = pool.Submit(func() {
			defer wg.Done()
			root.asyncScanRow(rows[index], ch, respParams...)
		})
	}

	go func(wg *sync.WaitGroup, ch chan *AsyncScanExRes, wp *ants.Pool) {
		wg.Wait()
		close(ch)
		wp.Release()
	}(&wg, ch, pool)

	return ch
}

func (root *Importer) IsHeaderConsistent(resps ...interface{}) (isConsistent bool, err error) {
	defer func() {
		if p := recover(); p != nil {
			err = fmt.Errorf("IsHeaderConsistent: internal error: %v", p)
		}
	}()
	if len(root.leafNodes) == 0 {
		root.leafNodes = root.getLeafNodes()
	}

	lastRespLength := 0
	for _, resp := range resps {
		v := reflect.ValueOf(resp).Elem()
		fieldNum := reflect.Indirect(v).NumField()
		for i := 0; i < fieldNum; i++ {
			field := reflect.Indirect(v).Type().Field(i)
			tag := field.Tag.Get("ex")
			path := strings.Split(tag, "|")

			leafNode := root.leafNodes[lastRespLength+i]
			if !reflect.DeepEqual(leafNode.path, path) {
				return
			}
		}

		lastRespLength += fieldNum
	}

	if lastRespLength != len(root.leafNodes) {
		return
	}

	isConsistent = true
	return
}
