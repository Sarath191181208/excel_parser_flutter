import 'package:excel/excel.dart';
import 'package:tuple/tuple.dart';

import 'cell_exceptions.dart';
import 'excel_data_classes.dart';

CellErrorsWarnings checkRepeatedElements(
    CellIndex headerIndex, List<Data?> headerColumn, String fieldName) {
  List<CellError> errors = [];
  List<CellWarning> warnings = [];
  Set<String> colValues = {};

  for (int rowIndex = 1; rowIndex < headerColumn.length; rowIndex++) {
    Data? name = headerColumn[rowIndex];
    CellIndex cellIndex = CellIndex.indexByColumnRow(
        columnIndex: headerIndex.columnIndex, rowIndex: rowIndex);
    if (name == null || name.value == null) {
      continue;
    }
    String dataString = name.value.toString();
    if (colValues.contains(dataString)) {
      errors.add(CellError('$fieldName is repeated: $dataString', cellIndex));
    } else {
      colValues.add(dataString);
    }
  }

  return Tuple2(errors, warnings);
}

CellErrorsWarnings checkEmptyElements(
    CellIndex headerIndex, List<Data?> headerColumn, String fieldName) {
  List<CellError> errors = [];
  List<CellWarning> warnings = [];

  for (int rowIndex = 1; rowIndex < headerColumn.length; rowIndex++) {
    Data? name = headerColumn[rowIndex];
    CellIndex cellIndex = CellIndex.indexByColumnRow(
        columnIndex: headerIndex.columnIndex, rowIndex: rowIndex);
    // check if the field data is empty
    if (name == null || name.value == null || name.value.toString().isEmpty) {
      warnings.add(CellWarning('$fieldName is empty', cellIndex));
    }
  }

  return Tuple2(errors, warnings);
}

CellErrorsWarnings checkIfElementsInList(CellIndex headerIndex,
    List<Data?> headerColumn, String fieldName, List<String> list) {
  List<CellError> errors = [];
  List<CellWarning> warnings = [];
  for (int rowIndex = 1; rowIndex < headerColumn.length; rowIndex++) {
    Data? name = headerColumn[rowIndex];
    CellIndex cellIndex = CellIndex.indexByColumnRow(
        columnIndex: headerIndex.columnIndex, rowIndex: rowIndex);
    if (name == null || name.value == null) {
      continue;
    }
    String dataString = name.value.toString();
    if (!list.contains(dataString)) {
      errors.add(
          CellError('$dataString is not defined for $fieldName', cellIndex));
    }
  }
  return Tuple2(errors, warnings);
}
