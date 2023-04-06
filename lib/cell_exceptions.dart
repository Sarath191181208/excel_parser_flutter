import 'package:excel/excel.dart';

class CellException {
  final String message;
  final CellIndex cellIndex;
  CellException(this.message, this.cellIndex);

  @override
  String toString() {
    return 'CellError{message: $message, cellIndex: $cellIndex}';
  }
}

class CellError extends CellException {
  CellError(String message, CellIndex cellIndex) : super(message, cellIndex);
}

class CellWarning extends CellException {
  CellWarning(String message, CellIndex cellIndex) : super(message, cellIndex);
}
