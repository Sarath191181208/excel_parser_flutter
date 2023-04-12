import 'package:excel/excel.dart';
import 'package:tuple/tuple.dart';

import 'cell_exceptions.dart';

var nameHeader = HeaderName(
    "Name", ['name', 'employee name', 'emp name', 'employee', 'full name']);
var emailHeader = HeaderName(
    "Email", ['email', 'phone', 'email/phone', 'email id / phone no']);
var departmentHeader =
    HeaderName("Department", ['department', 'dept', 'department name']);
var employeeHeader =
    HeaderName("Employee Id", ['employee id', 'emp id', 'empid']);

List<HeaderName> compulsaryHeaders = [
  nameHeader,
  emailHeader,
  departmentHeader,
  employeeHeader
];

typedef CellErrorsWarnings = Tuple2<List<CellError>, List<CellWarning>>;
typedef ValidatorFunction = CellErrorsWarnings Function(
    CellIndex, List<Data?>, String);

class CellNameAndIndexMap {
  late Map<String, CellIndex> map;
  CellNameAndIndexMap(Map<String, CellIndex> map) {
    this.map =
        map.map((key, value) => MapEntry(key.toLowerCase().trim(), value));
  }

  CellIndex? findHeader(RowHeaderComparable headerNames) {
    for (String key in map.keys) {
      if (headerNames.containsRowHeader(key)) {
        return map[key];
      }
    }
    return null;
  }

  bool contains(RowHeaderComparable headerNames) {
    return findHeader(headerNames) != null;
  }
}

abstract class RowHeaderComparable {
  bool containsRowHeader(String header);
  String get fieldName;
}

class HeaderName implements RowHeaderComparable {
  @override
  final String fieldName;
  final List<String> headers = [];
  HeaderName(this.fieldName, List<String> headers) {
    this.headers.addAll(headers.map(((e) => e.toLowerCase().trim())));
  }

  @override
  bool containsRowHeader(String header) {
    return headers.contains(header);
  }
}

class TimeHeaderName implements RowHeaderComparable {
  @override
  final String fieldName;

  TimeHeaderName(this.fieldName) {
    // check if the header is of format Shift{1-20} [start|end]
    RegExp regExp = RegExp(r"shift(\d{1,4}) (start|end)");
    var match = regExp.firstMatch(fieldName.toLowerCase());
    if (match == null) {
      throw Exception(
          "Invalid header name: $fieldName. Expected format: Shift{1-20} [start|end]");
    }
  }

  bool get isStartHeader => fieldName.toLowerCase().contains("start");
  int get shiftNumber {
    RegExp regExp = RegExp(r"shift(\d{1,4}) (start|end)");
    var match = regExp.firstMatch(fieldName.toLowerCase());
    if (match == null) {
      throw Exception(
          "Invalid header name: $fieldName. Expected format: Shift{1-20} [start|end]");
    }
    return int.parse(match.group(1)!);
  }

  @override
  bool containsRowHeader(String header) {
    return header.toLowerCase() == fieldName.toLowerCase();
  }
}
