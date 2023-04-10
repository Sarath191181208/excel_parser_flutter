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

  List<CellError> checkSheetHeaders(Sheet sheet,
      List<RowHeaderComaparable> headers, CellNameAndIndexMap map) {
    List<CellError> errors = [];
    for (RowHeaderComaparable header in headers) {
      CellIndex? index = map.findHeader(header);
      if (index == null) {
        errors.add(CellError('Header "${header.fieldName}" not found',
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
    CellNameAndIndexMap map = res.item1;
    List<CellError> headerErrors = res.item2;

    if (headerErrors.isNotEmpty) {
      log("Errors in headers: ");
      log(headerErrors.toString());
      return Tuple2(headerErrors, []);
    }

    // validating if all the headers are defined
    List<CellError> areAllHeadersDefined = checkSheetHeaders(
        sheet2, [...compulsaryHeaders, ...widget.shiftHeaders], map);

    if (areAllHeadersDefined.isNotEmpty) {
      log("Errors in are, All headers defined: ");
      log(areAllHeadersDefined.toString());
      return Tuple2(areAllHeadersDefined, []);
    }

    List<ValidatorFunction> validators = [
      validateRepetitionsForRow,
      validateEmptyElementsForRow,
    ];
    List<ValidatorFunction> deptValidators = [
      (RowHeaderComaparable h, Sheet s, CellNameAndIndexMap map) =>
          areRowItemsinList(h, s, map, list: widget.departments)
    ];

    // check if the name column has any repeated elements and empty elements
    var nameErrorsAndWarnings =
        applyValidators(nameHeader, sheet2, map, validators);
    var emailErrorsAndWarnings =
        applyValidators(emailHeader, sheet2, map, validators);
    var employeeErrosAndWarnings =
        applyValidators(employeeHeader, sheet2, map, validators);
    var departmentErrorsAndWarnings =
        applyValidators(departmentHeader, sheet2, map, deptValidators);

    showErrorsAndWarnings(nameErrorsAndWarnings, 'Name');
    showErrorsAndWarnings(emailErrorsAndWarnings, 'Email/Phone');
    showErrorsAndWarnings(departmentErrorsAndWarnings, 'Department');

    return mergeErrorsAndWarnings([
      nameErrorsAndWarnings,
      emailErrorsAndWarnings,
      employeeErrosAndWarnings,
      departmentErrorsAndWarnings,
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
