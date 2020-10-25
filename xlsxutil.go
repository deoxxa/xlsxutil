package xlsxutil

import (
	"encoding"
	"fmt"
	"reflect"
	"regexp"
	"strconv"
	"strings"

	"github.com/pkg/errors"
	"github.com/tealeg/xlsx"
)

func Sheet(doc *xlsx.File, name string) (*xlsx.Sheet, error) {
	var found []string

	for _, s := range doc.Sheets {
		found = append(found, s.Name)

		if Fuzzy(s.Name, name) {
			return s, nil
		}
	}

	return nil, errors.Errorf("Sheet: couldn't find sheet %q; options were %#v", name, found)
}

func Fuzzy(a, b string) bool {
	return strings.TrimSpace(strings.ToLower(a)) == strings.TrimSpace(strings.ToLower(b))
}

func CopyStyles(to, from *xlsx.Cell) {
	s1 := from.GetStyle()
	s2 := *s1
	to.SetStyle(&s2)
}

func Cell(r *xlsx.Row, n int) *xlsx.Cell {
	if len(r.Cells) == 0 {
		r.AddCell()
	}

	for i := len(r.Cells); i < n+1; i++ {
		CopyStyles(r.AddCell(), r.Cells[i-1])
	}

	return r.Cells[n]
}

type Money float64

func MoneyPointer(v Money) *Money { return &v }

func (m *Money) ScanString(s string) error {
	s = strings.Replace(s, "$", "", -1)
	s = strings.Replace(s, ",", "", -1)
	s = strings.Replace(s, " ", "", -1)
	s = strings.TrimSpace(s)

	if s == "" {
		*m = 0
		return nil
	}

	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return errors.Wrap(err, "Money.ScanString")
	}

	*m = Money(f)

	return nil
}

func (m Money) String() string {
	return fmt.Sprintf("$%.02f", float64(m))
}

func (m Money) Code() string {
	return fmt.Sprintf("%v", float64(m))
}

func (m Money) Round() string {
	return fmt.Sprintf("%.02f", float64(m))
}

func (m *Money) Valid() bool {
	return m != nil
}

type Years int

func YearsPointer(v Years) *Years { return &v }

func (y *Years) ScanString(s string) error {
	s = strings.ToLower(s)
	s = strings.TrimSuffix(s, "years")
	s = strings.TrimSuffix(s, "y")
	s = strings.Trim(s, "\t -")

	n, err := strconv.ParseInt(s, 10, 64)
	if err != nil {
		return errors.Wrap(err, "Years.ScanString")
	}

	*y = Years(n)

	return nil
}

func (y Years) String() string {
	return fmt.Sprintf("%d years", y)
}

func (y Years) Code() string {
	return fmt.Sprintf("%d", y)
}

func (y Years) Enum() string {
	return fmt.Sprintf("%d-years", y)
}

func (y Years) Months() Months {
	return Months(y * 12)
}

type Months int

func MonthsPointer(v Months) *Months { return &v }

func (m *Months) ScanString(s string) error {
	s = strings.ToLower(s)
	s = strings.TrimSuffix(s, "months")
	s = strings.TrimSuffix(s, "m")
	s = strings.Trim(s, "\t -")

	if s == "" {
		*m = 0
		return nil
	}

	n, err := strconv.ParseInt(s, 10, 64)
	if err != nil {
		return errors.Wrap(err, "Months.ScanString")
	}

	*m = Months(n)

	return nil
}

func (m Months) String() string {
	return fmt.Sprintf("%d months", m)
}

func (m Months) Code() string {
	return fmt.Sprintf("%d", m)
}

func (m Months) Enum() string {
	return fmt.Sprintf("%d-months", m)
}

type YesNo bool

func YesNoPointer(v YesNo) *YesNo { return &v }

func (y *YesNo) ScanString(s string) error {
	switch strings.ToLower(s) {
	case "yes", "y", "true":
		*y = true
	case "no", "n", "false", "":
		*y = false
	default:
		return fmt.Errorf("can't scan %q into YesNo", s)
	}

	return nil
}

func (y YesNo) String() string {
	if y {
		return "yes"
	}

	return "no"
}

func (y YesNo) Code() string {
	if y {
		return "true"
	}

	return "false"
}

type Range [2]int

func (r *Range) ScanString(s string) error {
	a := strings.Split(regexp.MustCompile("[^0-9]+").ReplaceAllString(s, " "), " ")

	if len(a) != 2 {
		return errors.Errorf("Range.ScanString: expected two components; instead got %d (%v from %q)", len(a), a, s)
	}

	n1, err := strconv.ParseInt(a[0], 10, 64)
	if err != nil {
		return err
	}

	n2, err := strconv.ParseInt(a[1], 10, 64)
	if err != nil {
		return err
	}

	r[0] = int(n1)
	r[1] = int(n2)

	return nil
}

func (r *Range) A() int { return r[0] }
func (r *Range) B() int { return r[1] }

func (r *Range) String() string {
	return fmt.Sprintf("%d-%d", r[0], r[1])
}

type Scanner interface {
	ScanString(s string) error
}

func Find(r *xlsx.Row, names ...string) map[string]int {
	res := make(map[string]int)

	for i, c := range r.Cells {
		for _, name := range names {
			if _, ok := res[name]; ok {
				continue
			}

			if Fuzzy(c.String(), name) {
				res[name] = i
			}
		}
	}

	return res
}

func FindHeader(s *xlsx.Sheet, limit int, names ...string) (int, map[string]int) {
	if limit >= len(s.Rows) {
		limit = len(s.Rows) - 1
	}

	bestRow := -1
	var bestCols map[string]int

	for i := 0; i <= limit; i++ {
		a := Find(s.Rows[i], names...)

		if len(a) == len(names) {
			return i, a
		}

		if len(a) > len(bestCols) {
			bestRow = i
			bestCols = a
		}
	}

	return bestRow, bestCols
}

func Scan(r *xlsx.Row, out ...interface{}) error {
	for i, e := range out {
		c := ""

		if len(r.Cells) > i {
			c = strings.TrimSpace(r.Cells[i].Value)
		}

		switch e := e.(type) {
		case nil:
			// nothing
		case *string:
			*e = c
		case *int:
			n, err := strconv.ParseInt(c, 10, 64)
			if err != nil {
				return errors.Wrapf(err, "Scan(%T)", e)
			}
			*e = int(n)
		case **int:
			if c == "" {
				*e = nil
			} else {
				n, err := strconv.ParseInt(c, 10, 64)
				if err != nil {
					return errors.Wrapf(err, "Scan(%T)", e)
				}
				v := int(n)
				*e = &v
			}
		case *float64:
			n, err := strconv.ParseFloat(c, 64)
			if err != nil {
				return errors.Wrapf(err, "Scan(%T)", e)
			}
			*e = n
		case **float64:
			if c == "" {
				*e = nil
			} else {
				n, err := strconv.ParseFloat(c, 64)
				if err != nil {
					return errors.Wrapf(err, "Scan(%T)", e)
				}
				*e = &n
			}
		default:
			p := reflect.ValueOf(e)

			if p.Type().Kind() != reflect.Ptr {
				return fmt.Errorf("can't scan into %T; must be a pointer", e)
			}

			if t := p.Type().Elem(); t.Kind() == reflect.Ptr && c == "" {
				p.Elem().Set(reflect.Zero(t))
				continue
			}

			if p.Type().Elem().Kind() == reflect.Ptr && p.Elem().IsNil() {
				p.Elem().Set(reflect.New(p.Type().Elem().Elem()))
				p = p.Elem()
			}

			v := p.Interface()

			if s, ok := v.(Scanner); ok {
				if err := s.ScanString(c); err != nil {
					return errors.Wrapf(err, "Scan(%T) (ScanString)", e)
				}

				continue
			}

			if s, ok := v.(encoding.TextUnmarshaler); ok {
				if err := s.UnmarshalText([]byte(c)); err != nil {
					return errors.Wrapf(err, "Scan(%T) (UnmarshalText)", e)
				}
			}

			return fmt.Errorf("can't scan into %T", e)
		}
	}

	return nil
}

func mapColumnNamesToFieldIndexes(t reflect.Type) ([]string, map[string]int) {
	a := make([]string, 0)
	m := make(map[string]int, 0)

	for i := 0; i < t.NumField(); i++ {
		f := t.Field(i)

		t, ok := f.Tag.Lookup("xlsx")
		if !ok {
			continue
		}

		n := strings.Split(t, ",")[0]

		a = append(a, n)

		m[n] = i
	}

	return a, m
}

type Adapter struct {
	s      *xlsx.Sheet
	typ    reflect.Type
	fields map[string]int
	cols   map[string]int
	width  int
	row    int
}

func newAdapter(s *xlsx.Sheet, typ reflect.Type) (*Adapter, error) {
	names, fields := mapColumnNamesToFieldIndexes(typ)
	if len(names) == 0 {
		return nil, errors.Errorf("newAdapter: couldn't find column names in struct tags")
	}

	row, cols := FindHeader(s, 10, names...)
	if len(cols) != len(fields) {
		var missing []string

		for k := range fields {
			if _, ok := cols[k]; !ok {
				missing = append(missing, k)
			}
		}

		return nil, errors.Errorf("newAdapter: couldn't find some required columns: %s", strings.Join(missing, ", "))
	}

	var width int
	for _, c := range cols {
		if c > width {
			width = c
		}
	}

	return &Adapter{
		s:      s,
		typ:    typ,
		fields: fields,
		cols:   cols,
		width:  width,
		row:    row,
	}, nil
}

func NewAdapter(s *xlsx.Sheet, v interface{}) (*Adapter, error) {
	a, err := newAdapter(s, reflect.TypeOf(v))
	if err != nil {
		return nil, errors.Wrap(err, "NewAdapter")
	}

	return a, nil
}

func newAdapterForSheet(doc *xlsx.File, name string, typ reflect.Type) (*Adapter, error) {
	s, err := Sheet(doc, name)
	if err != nil {
		return nil, errors.Wrap(err, "newAdapterForSheet")
	}

	return newAdapter(s, typ)
}

func NewAdapterForSheet(doc *xlsx.File, name string, v interface{}) (*Adapter, error) {
	a, err := newAdapterForSheet(doc, name, reflect.TypeOf(v))
	if err != nil {
		return nil, errors.Wrap(err, "NewAdapterForSheet")
	}

	return a, nil
}

func (r *Adapter) Next() bool {
	if r.row >= len(r.s.Rows)-1 {
		return false
	}

	r.row++

	for _, c := range r.s.Rows[r.row].Cells {
		if c.String() != "" {
			return true
		}
	}

	return r.Next()
}

func (r *Adapter) Read(out interface{}) error {
	p := reflect.ValueOf(out)
	if typ := reflect.PtrTo(r.typ); p.Type() != typ {
		return errors.Errorf("Adapter.Read: expected out to be %s; was instead %s", typ, p.Type())
	}

	arr := make([]interface{}, r.width+1)

	v := p.Elem()

	for name, f := range r.fields {
		arr[r.cols[name]] = v.Field(f).Addr().Interface()
	}

	if err := Scan(r.s.Rows[r.row], arr...); err != nil {
		return errors.Wrapf(err, "Adapter.Read: couldn't read row %d of %d", r.row, len(r.s.Rows))
	}

	return nil
}

func ReadAll(doc *xlsx.File, name string, out interface{}) error {
	p := reflect.ValueOf(out)
	if p.Kind() != reflect.Ptr {
		return errors.Errorf("ReadAll: expected out to be pointer; was instead %s", p.Kind())
	}

	s := p.Elem()
	if s.Kind() != reflect.Slice {
		return errors.Errorf("ReadAll: expected out to be pointer to slice; was instead pointer to %s", s.Kind())
	}

	t := s.Type().Elem()
	if t.Kind() != reflect.Struct {
		return errors.Errorf("ReadAll: expected out to be pointer to slice of struct; was instead pointer to slice of %s", t.Kind())
	}

	rd, err := newAdapterForSheet(doc, name, t)
	if err != nil {
		return errors.Wrap(err, "ReadAll: couldn't construct adapter")
	}

	for rd.Next() {
		e := reflect.New(t)

		if err := rd.Read(e.Interface()); err != nil {
			return errors.Wrapf(err, "ReadAll: couldn't read row %d of %d", rd.row, len(rd.s.Rows))
		}

		s.Set(reflect.Append(s, reflect.Indirect(e)))
	}

	return nil
}

func (r *Adapter) Write(in interface{}) error {
	p := reflect.ValueOf(in)
	if p.Type() != r.typ {
		return errors.Errorf("Adapter.Write: expected in to be %s; was instead %s", r.typ, p.Type())
	}

	for name, f := range r.fields {
		v := p.Field(f)
		e := v.Interface()

		switch e := e.(type) {
		case nil:
			Cell(r.s.Rows[r.row], r.cols[name]).SetString("")
		case string:
			Cell(r.s.Rows[r.row], r.cols[name]).SetString(e)
		case *string:
			if e == nil {
				Cell(r.s.Rows[r.row], r.cols[name]).SetString("")
			} else {
				Cell(r.s.Rows[r.row], r.cols[name]).SetString(*e)
			}
		case float64:
			Cell(r.s.Rows[r.row], r.cols[name]).SetString(fmt.Sprintf("%v", e))
		case interface{ Enum() string }:
			if v.Kind() == reflect.Ptr && v.IsNil() {
				Cell(r.s.Rows[r.row], r.cols[name]).SetString("")
			} else {
				Cell(r.s.Rows[r.row], r.cols[name]).SetString(e.Enum())
			}
		case fmt.Stringer:
			if v.Kind() == reflect.Ptr && v.IsNil() {
				Cell(r.s.Rows[r.row], r.cols[name]).SetString("")
			} else {
				Cell(r.s.Rows[r.row], r.cols[name]).SetString(e.String())
			}
		default:
			return errors.Errorf("Adapter.Write: can't write field of type %T", e)
		}
	}

	return nil
}

func WriteAll(doc *xlsx.File, name string, in interface{}) error {
	p := reflect.ValueOf(in)
	if p.Kind() != reflect.Slice {
		return errors.Errorf("WriteAll: expected in to be slice; was instead %s", p.Kind())
	}

	t := p.Type().Elem()
	if t.Kind() != reflect.Struct {
		return errors.Errorf("WriteAll: expected in to be slice of struct; was instead slice of %s", t.Kind())
	}

	ad, err := newAdapterForSheet(doc, name, t)
	if err != nil {
		return errors.Wrap(err, "WriteAll: couldn't construct adapter")
	}

	ad.s.Rows = ad.s.Rows[0 : ad.row+1]

	for i, j := 0, p.Len(); i < j; i++ {
		ad.s.AddRow()

		ad.Next()

		if err := ad.Write(p.Index(i).Interface()); err != nil {
			return errors.Wrapf(err, "WriteAll: couldn't write row %d of %d", ad.row, j)
		}
	}

	return nil
}

func SetupSheet(doc *xlsx.File, name string, in interface{}) (*xlsx.Sheet, error) {
	res, err := setupSheet(doc, name, reflect.TypeOf(in))
	if err != nil {
		return nil, errors.Wrap(err, "SetupSheet")
	}

	return res, nil
}

func setupSheet(doc *xlsx.File, name string, t reflect.Type) (*xlsx.Sheet, error) {
	names, fields := mapColumnNamesToFieldIndexes(t)
	if len(names) == 0 {
		return nil, errors.Errorf("setupSheet: couldn't find column names in struct tags")
	}

	s, err := Sheet(doc, name)
	if err != nil {
		ss, err := doc.AddSheet(name)
		if err != nil {
			return nil, errors.Wrap(err, "setupSheet: couldn't add sheet")
		}

		s = ss
	}

	if _, cols := FindHeader(s, 10, names...); len(cols) == len(fields) {
		return s, nil
	}

	r := s.AddRow()

	for _, v := range names {
		r.AddCell().SetString(v)
	}

	return s, nil
}

func SetupSheetAndWriteAll(doc *xlsx.File, name string, in interface{}) error {
	p := reflect.ValueOf(in)
	if p.Kind() != reflect.Slice {
		return errors.Errorf("SetupSheetAndWriteAll: expected in to be slice; was instead %s", p.Kind())
	}

	t := p.Type().Elem()
	if t.Kind() != reflect.Struct {
		return errors.Errorf("SetupSheetAndWriteAll: expected in to be slice of struct; was instead slice of %s", t.Kind())
	}

	if _, err := setupSheet(doc, name, t); err != nil {
		return errors.Wrap(err, "SetupSheetAndWriteAll: couldn't run setupSheet")
	}

	if err := WriteAll(doc, name, in); err != nil {
		return errors.Wrap(err, "SetupSheetAndWriteAll: couldn't run WriteAll")
	}

	return nil
}
