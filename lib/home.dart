import 'dart:io';

import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:open_file/open_file.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart';
import 'package:path_provider/path_provider.dart';

class ExcelTable extends StatefulWidget {
  @override
  State<ExcelTable> createState() => _ExcelTableState();
}

class _ExcelTableState extends State<ExcelTable> {

Future<void> createExcel() async {
  FilePickerResult? result = await FilePicker.platform.pickFiles(
    type: FileType.custom,
    allowedExtensions: ['xlsx'],
  );

  final Workbook workbook = Workbook();
  final Worksheet sheet = workbook.worksheets[5];

  // Set the values for the table
  sheet.getRangeByName('A').setText('OrderDate');
  sheet.getRangeByName('B').setText('Region');
  sheet.getRangeByName('C').setText('Rep');
  sheet.getRangeByName('D').setText('Item');
  sheet.getRangeByName('E').setText('Units cost');
  sheet.getRangeByName('F').setText('unit');
  sheet.getRangeByName('G').setText('Total');

  // Create a table with the values
  // Save the workbook
  final List<int> bytes = workbook.saveAsStream();
  workbook.dispose();
  final String path = (await getApplicationSupportDirectory()).path;
  final String fileName = '$path/Output.xlsx';

  final File file = File(fileName);
  await file.writeAsBytes(bytes, flush: true);
  OpenFile.open(fileName);// Save the bytes to a file or send them to a server
}



  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: InteractiveViewer(
        child: Table(
          border: TableBorder.all(),
          children: [
            TableRow(
              children: [
                TableCell(child: Text('OrderDate'),),
                TableCell(child: Text('Region')),
                TableCell(child: Text('Rep')),
                TableCell(child: Text('Item')),
                TableCell(child: Text('Units'))
                ,TableCell(child: Text('Unit cost'))
                ,TableCell(child: Text('Total'))
                ,

              ],
            ),
            TableRow(
              children: [
                TableCell(child: Text('')),
                TableCell(child: Text('')),
                TableCell(child: Text(''))
                ,TableCell(child: Text(''))
                ,TableCell(child: Text(''))
                ,TableCell(child: Text(''))
                ,TableCell(child: Text(''))
                ,
              ],
            ),
            TableRow(
              children: [
                TableCell(child: Text('')),
                TableCell(child: Text('')),
                TableCell(child: Text('')),
                TableCell(child: Text('')),
                TableCell(child: Text('')),
                TableCell(child: Text('')),
                TableCell(child: Text('')),
              ],
            ),
          ],
        ),
      ),
      floatingActionButton: FloatingActionButton(
        onPressed: createExcel,
        tooltip: 'Open Excel file',
        child: Icon(Icons.folder_open),
      ),
    );
  }
}