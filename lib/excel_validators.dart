import 'dart:developer';

import 'package:excel/excel.dart';
import 'package:intl/intl.dart';
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

CellErrorsWarnings checkTimeFormat(
    CellIndex headerIndex, List<Data?> headerColumn, String fieldName) {
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
    if (dataString.isEmpty) {
      continue;
    }
    try {
      dataString = dataString.trim();
      dataString = dataString.replaceAll(RegExp(r'\s*:\s*'), ':');
      dataString = dataString.toLowerCase();
      parseTime(dataString).toString();
    } catch (e) {
      log(e.toString());
      errors.add(CellError(
          '$dataString is not a valid time it must be of the form 12:00 AM',
          cellIndex));
    }
  }
  return Tuple2(errors, warnings);
}

CellErrorsWarnings compareStartEndTimes(
    CellIndex headerIndex, List<Data?> startColumn, String fieldName,
    {required List<Data?> endColumn, required String endFieldName}) {
  List<CellError> errors = [];
  List<CellWarning> warnings = [];
  for (int rowIndex = 1; rowIndex < startColumn.length; rowIndex++) {
    Data? name = startColumn[rowIndex];
    CellIndex cellIndex = CellIndex.indexByColumnRow(
        columnIndex: headerIndex.columnIndex, rowIndex: rowIndex);
    if (name == null || name.value == null) {
      warnings.add(CellWarning('$fieldName is empty', cellIndex));
      continue;
    }
    String dataString = name.value.toString();
    if (dataString.isEmpty) {
      warnings.add(CellWarning('$fieldName is empty', cellIndex));
      continue;
    }
    DateTime? startTime = tryParseTime(dataString);
    if (startTime == null) {
      errors.add(CellError(
          '$dataString is not a valid time it must be of the form 12:00 AM',
          cellIndex));
      continue;
    }
    Data? endTimeHeader = endColumn[rowIndex];
    if (endTimeHeader == null || endTimeHeader.value == null) {
      warnings.add(CellWarning('$endFieldName is empty', cellIndex));
      continue;
    }
    String endTimeString = endTimeHeader.value.toString();
    if (endTimeString.isEmpty) {
      warnings.add(CellWarning('$endFieldName is empty', cellIndex));
      continue;
    }
    DateTime? endTime = tryParseTime(endTimeString);
    if (endTime == null) {
      errors.add(CellError(
          '$endTimeString is not a valid time it must be of the form 12:00 AM',
          cellIndex));
      continue;
    }

    if (endTime.isBefore(startTime)) {
      errors
          .add(CellError('$endFieldName must be after $fieldName', cellIndex));
    }
  }
  return Tuple2(errors, warnings);
}

DateTime? tryParseTime(String time) {
  try {
    time = time.trim();
    time = time.replaceAll(RegExp(r'\s*:\s*'), ':');
    time = time.toLowerCase();
    return parseTime(time);
  } catch (e) {
    return null;
  }
}

DateTime parseTime(String time) {
  /// This function parses a time string and returns a DateTime object
  /// The time string can be of the form 12:00 AM or 12:00:00 AM
  /// If the time string is invalid an exception is thrown

  List<String> formats = [
    'h:mm a',
    'hh:mm a',
    'h:mm:ss a',
    'hh:mm:ss a'
        'H:mm',
    'HH:mm',
    'H:mm:ss',
    'HH:mm:ss'
  ];
  for (String format in formats) {
    try {
      DateFormat formatter = DateFormat(format);
      return formatter.parse(time);
    } catch (e) {
      log(e.toString());
    }
  }
  return DateTime.parse(time.toUpperCase());
}
