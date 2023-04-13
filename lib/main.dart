// ignore_for_file: public_member_api_docs, sort_constructors_first
import 'dart:developer';
import 'dart:io';
import 'dart:typed_data';

import 'package:downloads_path_provider_28/downloads_path_provider_28.dart';
import 'package:excel/excel.dart';
import 'package:excel_parser/utils.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:tuple/tuple.dart';

import 'cell_exceptions.dart';
import 'excel_data_classes.dart';
import 'excel_logic.dart';
import 'excel_validators.dart';

void main() {
  runApp(const MyApp());
}

Future<void> clearTemporaryFiles() async {
  Directory dir = await getTemporaryDirectory();
  dir.deleteSync(recursive: true);
  dir.create(); // This will create the temporary directory again. So temporary files will only be deleted
}

List<TimeHeaderName> buildHeaderFromStr(List<String> name) {
  return name.map((e) => TimeHeaderName(e)).toList();
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
        title: 'Flutter Demo',
        theme: ThemeData(
          primarySwatch: Colors.blue,
        ),
        home: MyHomePage(
            title: 'Flutter Demo Home Page',
            departments: const [
              'IT',
              'HR',
              'Finance',
              'Marketing',
              'Sales',
              'Admin'
            ],
            shiftHeaders: buildHeaderFromStr([
              "Shift1 Start",
              "Shift1 End",
              "Shift2 Start",
              "Shift2 End",
            ])));
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage(
      {super.key,
      required this.title,
      required this.departments,
      required this.shiftHeaders});

  final String title;
  final List<String> departments;
  final List<TimeHeaderName> shiftHeaders;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  Future<void> testExcelSave() async {
    Excel excel = Excel.createExcel();
    Sheet sheet = excel['Sheet1'];

    // create column names like name, email/phone, department etc
    createColumnNames(sheet);

    // save the file to a path chosen by the user
    var dir = await DownloadsPathProvider.downloadsDirectory;
    List<int>? bytes = excel.save();
    if (bytes == null || bytes.isEmpty || dir == null) {
      return;
    }

    Uint8List data = Uint8List.fromList(bytes);
    writeFileToDownloads(data, dir, 'test.xlsx');
  }

  List<CellError> checkSheetHeaders(
      Sheet sheet, List<RowHeaderComparable> headers, CellNameAndIndexMap map) {
    List<CellError> errors = [];
    for (RowHeaderComparable header in headers) {
      CellIndex? index = map.findHeader(header);
      if (index == null) {
        errors.add(CellError('Header "${header.headerName}" not found',
            CellIndex.indexByString('A1')));
      }
    }
    return errors;
  }

  CellErrorsWarnings? listErrorsWarnings;

  Future<CellErrorsWarnings> parseExcelFile(String path) async {
    // select and read the excel sheet
    var file = File(path);
    Uint8List readBytes = file.readAsBytesSync();
    Excel excel2 = Excel.decodeBytes(readBytes);
    Sheet sheet2 = excel2['Sheet1'];

    // get the cell indices of the names column
    var res = getColumnIdsFromNames(sheet2);
    CellNameAndIndexMap rowHeaderCellIndexMap = res.item1;
    List<CellError> headerErrors = res.item2;

    if (headerErrors.isNotEmpty) {
      log("Errors in headers: ");
      log(headerErrors.toString());
      return Tuple2(headerErrors, []);
    }

    // validating if all the headers are defined
    List<CellError> areAllHeadersDefined = checkSheetHeaders(sheet2,
        [...compulsaryHeaders, ...widget.shiftHeaders], rowHeaderCellIndexMap);

    if (areAllHeadersDefined.isNotEmpty) {
      log("Errors in are, All headers defined: ");
      log(areAllHeadersDefined.toString());
      return Tuple2(areAllHeadersDefined, []);
    }

    List<ValidatorFunction> validators = [
      checkRepeatedElements,
      checkEmptyElements,
    ];
    List<ValidatorFunction> deptValidators = [
      (CellIndex c, List<Data?> d, String name) =>
          checkIfElementsInList(c, d, name, widget.departments)
    ];

    List<ValidatorFunction> shiftValidators = [
      checkEmptyElements,
      checkTimeFormat
    ];

    runValidator(RowHeaderComparable h, List<ValidatorFunction> v) =>
        applyValidators(h, sheet2, rowHeaderCellIndexMap, v);

    // check if the name column has any repeated elements and empty elements
    var nameErrorsAndWarnings = runValidator(nameHeader, validators);
    var emailErrorsAndWarnings = runValidator(emailHeader, validators);
    var employeeErrosAndWarnings = runValidator(employeeHeader, validators);
    var departmentErrorsAndWarnings =
        runValidator(departmentHeader, deptValidators);

    List<CellErrorsWarnings> shiftErrorsAndWarnings = [];
    for (TimeHeaderName header in widget.shiftHeaders) {
      shiftErrorsAndWarnings.add(runValidator(header, shiftValidators));
    }

    var groupedShiftHeadersMap = getGroupedStartEndTimes(widget.shiftHeaders);

    List<CellError> errors = [];
    List<CellWarning> warnings = [];
    // check if the start time is before the end time
    for (int key in groupedShiftHeadersMap.keys) {
      StartEndTimesTuple tuple = groupedShiftHeadersMap[key]!;
      TimeHeaderName? startTimeHeader = tuple.item1;
      TimeHeaderName? endTimeHeader = tuple.item2;
      if (startTimeHeader == null) {
        errors.add(CellError('Start time header not found for shift $key',
            CellIndex.indexByString('A1')));
        continue;
      }
      if (endTimeHeader == null) {
        errors.add(CellError('End time header not found for shift $key',
            CellIndex.indexByString('A1')));
        continue;
      }

      var endColumnIndex = rowHeaderCellIndexMap.findHeader(endTimeHeader)!;
      List<Data?> endColumn =
          selectColumn(rowIndex: endColumnIndex, sheet: sheet2);

      List<ValidatorFunction> startEndValidators = [
        (CellIndex c, List<Data?> d, String name) => compareStartEndTimes(
            c, d, name,
            endColumn: endColumn, endFieldName: endTimeHeader.headerName),
      ];
      shiftErrorsAndWarnings
          .add(runValidator(startTimeHeader, startEndValidators));
    }

    shiftErrorsAndWarnings.add(Tuple2(errors, warnings));

    // showErrorsAndWarnings(nameErrorsAndWarnings, 'Name');
    // showErrorsAndWarnings(emailErrorsAndWarnings, 'Email/Phone');
    // showErrorsAndWarnings(departmentErrorsAndWarnings, 'Department');

    return mergeErrorsAndWarnings([
      nameErrorsAndWarnings,
      emailErrorsAndWarnings,
      employeeErrosAndWarnings,
      departmentErrorsAndWarnings,
      ...shiftErrorsAndWarnings
    ]);
  }

  Future<void> selectAndReadFile() async {
    // get the file to read
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );
    // This is required as the file picker caches the opened file by default and ther isn't a way to stop it
    clearTemporaryFiles();
    if (result == null) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(
          content: Text('No file selected'),
        ),
      );
      return;
    }

    String path = result.files.single.path!;

    CellErrorsWarnings exceptions = await parseExcelFile(path);
    setState(() {
      listErrorsWarnings = exceptions;
    });
  }

  @override
  Widget build(BuildContext context) {
    final screenHeight = MediaQuery.of(context).size.height;
    return Scaffold(
      appBar: AppBar(
        title: Text(widget.title),
      ),
      body: SafeArea(
        child: SingleChildScrollView(
          child: Column(
            mainAxisAlignment: MainAxisAlignment.center,
            children: [
              SizedBox(height: screenHeight * 0.20),
              ElevatedButton(
                onPressed: testExcelSave,
                child: const Text('Download Text Excel File'),
              ),
              ElevatedButton(
                onPressed: selectAndReadFile,
                child: const Text('Read Excel File'),
              ),
              if (listErrorsWarnings != null)
                ErrorsWarningsWidget(cellErrorsAndWarnings: listErrorsWarnings!)
            ],
          ),
        ),
      ),
    );
  }
}

class ErrorsWarningsWidget extends StatelessWidget {
  final CellErrorsWarnings cellErrorsAndWarnings;
  const ErrorsWarningsWidget({super.key, required this.cellErrorsAndWarnings});

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        for (CellError cellError in cellErrorsAndWarnings.item1)
          ListTile(
            leading: const Icon(Icons.error),
            title: Text(cellError.message),
            iconColor: Colors.red,
            trailing: Text(
                "row: ${cellError.cellIndex.rowIndex} col: ${cellError.cellIndex.columnIndex}"),
          ),
        for (CellWarning cellWarning in cellErrorsAndWarnings.item2)
          ListTile(
            leading: const Icon(Icons.error),
            title: Text(cellWarning.message),
            iconColor: Colors.yellow,
            trailing: Text(
                "row: ${cellWarning.cellIndex.rowIndex} col: ${cellWarning.cellIndex.columnIndex}"),
          ),
      ],
    );
  }
}
