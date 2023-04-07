import 'package:excel/excel.dart';
import 'package:tuple/tuple.dart';

import 'cell_exceptions.dart';

var nameHeader = HeaderNamesList(
    "Name", ['name', 'employee name', 'emp name', 'employee', 'full name']);
var emailHeader = HeaderNamesList(
    "Email", ['email', 'phone', 'email/phone', 'email id / phone no']);
var departmentHeader =
    HeaderNamesList("Department", ['department', 'dept', 'department name']);
var employeeHeader =
    HeaderNamesList("Employee Id", ['employee id', 'emp id', 'empid']);

List<HeaderNamesList> compulsaryHeaders = [
  nameHeader,
  emailHeader,
  departmentHeader,
  employeeHeader
];

typedef CellErrorsWarnings = Tuple2<List<CellError>, List<CellWarning>>;
typedef ValidatorFunction = CellErrorsWarnings Function(
    HeaderNamesList, Sheet, CellNameAndIndexMap);

class CellNameAndIndexMap {
  late Map<String, CellIndex> map;
  CellNameAndIndexMap(Map<String, CellIndex> map) {
    this.map =
        map.map((key, value) => MapEntry(key.toLowerCase().trim(), value));
  }

  CellIndex? findHeader(HeaderNamesList headerNames) {
    for (String header in headerNames.headers) {
      String headerName = header.toLowerCase().trim();
      if (map.containsKey(headerName)) {
        return map[header];
      }
    }
    return null;
  }

  bool contains(HeaderNamesList headerNames) {
    return findHeader(headerNames) != null;
  }
}

class HeaderNamesList {
  final String fieldName;
  final List<String> headers = [];
  HeaderNamesList(this.fieldName, List<String> headers) {
    this.headers.addAll(headers.map(((e) => e.toLowerCase().trim())));
  }

  bool contains(String header) {
    return headers.contains(header);
  }
}
