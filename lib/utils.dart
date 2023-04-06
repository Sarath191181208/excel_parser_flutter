import 'dart:developer';

import 'package:tuple/tuple.dart';
import 'cell_exceptions.dart';
import 'excel_data_classes.dart';

CellErrorsWarnings mergeErrorsAndWarnings(
    List<CellErrorsWarnings> errorsAndWarnings) {
  List<CellError> errors = [];
  List<CellWarning> warnings = [];
  for (CellErrorsWarnings errorsAndWarning in errorsAndWarnings) {
    errors.addAll(errorsAndWarning.item1);
    warnings.addAll(errorsAndWarning.item2);
  }
  return Tuple2(errors, warnings);
}

void showErrorsAndWarnings(CellErrorsWarnings errors, String fieldName) {
  log("Errors in $fieldName: ");
  log(errors.item1.toString());
  log("Warnings in $fieldName:  ");
  log(errors.item2.toString());
}
