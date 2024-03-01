import 'dart:io';

import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:path/path.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:open_filex/open_filex.dart';
// import "package:excel_example/dummyData.dart";
import "package:excel_example/model/member.dart";

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'Flutter Demo Home Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({Key? key, required this.title}) : super(key: key);

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  Future<void> createAndSaveExcel() async {
    var excel = Excel.createExcel();
    Sheet sheetObject = excel['Sheet1'];

    sheetObject.merge(
        CellIndex.indexByString('A1'), CellIndex.indexByString('I1'),
        customValue: TextCellValue('जमाखर्च पुस्तक'));
    sheetObject.merge(
        CellIndex.indexByString('A2'), CellIndex.indexByString('A3'),
        customValue: TextCellValue("अ. क्र."));
    sheetObject.merge(
        CellIndex.indexByString('B2'), CellIndex.indexByString('B3'),
        customValue: TextCellValue("सभासदाचे नाव"));

    sheetObject.merge(
        CellIndex.indexByString('C2'), CellIndex.indexByString('I2'),
        customValue: TextCellValue('जमा'));
    sheetObject.merge(
        CellIndex.indexByString('J2'), CellIndex.indexByString('L2'),
        customValue: TextCellValue('नावे'));

    sheetObject.merge(
        CellIndex.indexByString('M2'), CellIndex.indexByString('O2'),
        customValue: TextCellValue('आज अखेर बचत '));

    sheetObject.merge(
        CellIndex.indexByString('P2'), CellIndex.indexByString('S2'),
        customValue: TextCellValue('आज अखेर अर्थ सहाय'));

    sheetObject.cell(CellIndex.indexByString('C3')).value =
        TextCellValue('मासिक बचत');
    sheetObject.cell(CellIndex.indexByString('D3')).value =
        TextCellValue('जादा बचत');

    sheetObject.cell(CellIndex.indexByString('E3')).value =
        TextCellValue('आर्थिक सहाय परतफेड	');
    sheetObject.cell(CellIndex.indexByString('F3')).value =
        TextCellValue('सेवाशुल्क');
    sheetObject.cell(CellIndex.indexByString('G3')).value =
        TextCellValue('दंड');
    sheetObject.cell(CellIndex.indexByString('H3')).value =
        TextCellValue('इतर जमा');
    sheetObject.cell(CellIndex.indexByString('I3')).value =
        TextCellValue('एकूण');
    sheetObject.cell(CellIndex.indexByString('J3')).value =
        TextCellValue('दिलेले आर्थिक सहाय');
    sheetObject.cell(CellIndex.indexByString('K3')).value =
        TextCellValue('बचत परत व्याज');

    sheetObject.cell(CellIndex.indexByString('L3')).value =
        TextCellValue('एकूण');

    sheetObject.cell(CellIndex.indexByString('M3')).value =
        TextCellValue('एकूण बचत ');

    sheetObject.cell(CellIndex.indexByString('N3')).value =
        TextCellValue('जादा बचत ');

    sheetObject.cell(CellIndex.indexByString('O3')).value =
        TextCellValue('बचत परत');

    sheetObject.cell(CellIndex.indexByString('P3')).value =
        TextCellValue('घेतलेले आर्थिक सहाय');

    sheetObject.cell(CellIndex.indexByString('Q3')).value =
        TextCellValue('परतफेड');
    sheetObject.cell(CellIndex.indexByString('R3')).value =
        TextCellValue('सेवाशुल्क');
    sheetObject.cell(CellIndex.indexByString('S3')).value =
        TextCellValue('आर्थिक सहाय बाकी');

    sheetObject.merge(
        CellIndex.indexByString('T2'), CellIndex.indexByString("T3"),
        customValue: TextCellValue("दंड"));
    sheetObject.merge(
        CellIndex.indexByString('U2'), CellIndex.indexByString("U3"),
        customValue: TextCellValue("इतर जमा"));

    List<Member> dummyData = generateDummyData();
    for (int i = 0; i < dummyData.length; i++) {
      Member member = dummyData[i];
      List<String> rowData = [
        (i + 1).toString(),
        member.name,
        member.monthlySavings,
        member.extraSavings,
        member.loanRepaid,
        member.serviceCharge,
        member.penalty,
        member.otherDeposits,
        member.totalDeposit,
        member.givenLoan,
        member.loanInterest,
        member.totalLoanPaid,
        member.totalSavings,
        member.totalExtraSavings,
        member.savingsBack,
        member.loanTakenTillDate,
        member.totalLoanReturn,
        member.remainingLoan,
        member.totalServiceCharge,
        member.totalPenalty,
        member.totalOtherSavings,
      ];
      List<CellValue?> rowCells = rowData.map((cellData) {
        return (TextCellValue(cellData.toString()));
      }).toList();

      sheetObject.insertRowIterables(
          rowCells, 4 + i); // Start appending data after 4th row
    }

    var fileBytes = excel.save();
  }

  List<Member> generateDummyData() {
    List<Member> dummyData = [];

    for (int i = 1; i <= 20; i++) {
      Member member = Member(
        name: 'Person $i',
        monthlySavings: '100',
        extraSavings: '50',
        loanRepaid: '200',
        serviceCharge: '10',
        penalty: '5',
        otherDeposits: '50',
        totalDeposit: '500',
        givenLoan: '1000',
        loanInterest: '100',
        totalLoanPaid: '500',
        totalSavings: '1000',
        totalExtraSavings: '500',
        savingsBack: '200',
        loanTakenTillDate: '1500',
        totalLoanReturn: '1000',
        remainingLoan: '500',
        totalServiceCharge: '50',
        totalPenalty: '20',
        totalOtherSavings: '100',
      );

      dummyData.add(member);
    }

    return dummyData;
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: <Widget>[
            const Text(
              'You have pushed the button this many times:',
            ),
            ElevatedButton(
              onPressed: () {
                createAndSaveExcel();
              },
              child: const Text('Create Excel'),
            ),
          ],
        ),
      ),
    );
  }
}
