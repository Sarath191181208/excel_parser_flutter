import 'dart:io';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:excel_parser/cell_exceptions.dart';
import 'package:excel_parser/excel_validators.dart';
import 'package:path/path.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:tuple/tuple.dart';
import 'excel_data_classes.dart';

List<Data?> selectColumn({required CellIndex rowIndex, required Sheet sheet}) {
  List<Data?> rowData = [];
  for (var row in sheet.rows) {
    Data? data = row[rowIndex.columnIndex];
    rowData.add(data);
  }
  return rowData;
}

Tuple2<CellNameAndIndexMap, List<CellError>> getColumnIdsFromNames(
    Sheet sheet) {
  Map<String, CellIndex> columnIds = {};
  List<CellError> errors = [];

  for (Data? cell in sheet.rows[0]) {
    if (cell != null) {
      final headerName = cell.value.toString().trim();
      if (headerName.isEmpty) {
        errors.add(CellError('Empty string', cell.cellIndex));
      }
      if (columnIds.containsKey(headerName.toString())) {
        errors.add(CellError('Duplicate header $headerName', cell.cellIndex));
      }
      columnIds[headerName.toString()] = cell.cellIndex;
    }
  }
  return Tuple2.fromList([CellNameAndIndexMap(columnIds), errors]);
}

CellErrorsWarnings applyValidators(RowHeaderComparable header, Sheet sheet,
    CellNameAndIndexMap map, List<ValidatorFunction> validators) {
  List<CellError> errors = [];
  List<CellWarning> warnings = [];
  for (ValidatorFunction validator in validators) {
    CellIndex headerIndex = map.findHeader(header)!;
    List<Data?> headerColumn =
        selectColumn(rowIndex: headerIndex, sheet: sheet);
    CellErrorsWarnings result =
        validator(headerIndex, headerColumn, header.headerName);
    errors.addAll(result.item1);
    warnings.addAll(result.item2);
  }
  return Tuple2(errors, warnings);
}

Map<int, StartEndTimesTuple> getGroupedStartEndTimes(
    List<TimeHeaderName> shiftHeaders) {
  Map<int, StartEndTimesTuple> groupedShiftHeadersMap = {};
  for (TimeHeaderName header in shiftHeaders) {
    int shiftNumber = header.shiftNumber;
    if (!groupedShiftHeadersMap.containsKey(shiftNumber)) {
      groupedShiftHeadersMap[shiftNumber] = const Tuple2(null, null);
    }
    if (header.isStartHeader) {
      groupedShiftHeadersMap[shiftNumber] =
          Tuple2(header, groupedShiftHeadersMap[shiftNumber]!.item2);
    } else {
      groupedShiftHeadersMap[shiftNumber] =
          Tuple2(groupedShiftHeadersMap[shiftNumber]!.item1, header);
    }
  }
  return groupedShiftHeadersMap;
}

void createColumnNames(Sheet sheet) {
  // create a name column
  sheet.cell(CellIndex.indexByString('A1')).value = 'Name';

  // create a emp id column
  sheet.cell(CellIndex.indexByString('B1')).value = 'Emp Id';

  // create a emai id / phone no column
  sheet.cell(CellIndex.indexByString('C1')).value = 'Email Id / Phone No';

  // create a department column
  sheet.cell(CellIndex.indexByString('D1')).value = 'Department';
}

Future<File> writeFileToDownloads(
    Uint8List data, Directory root, String name) async {
  // storage permission ask
  var status = await Permission.storage.status;
  if (!status.isGranted) {
    await Permission.storage.request();
  }
  var filePath = join(root.path, name);
  print(filePath);
  // the data
  var bytes = ByteData.view(data.buffer);
  final buffer = bytes.buffer;
  // save the data in the path
  return File(filePath)
      .writeAsBytes(buffer.asUint8List(data.offsetInBytes, data.lengthInBytes));
}
