# ⚠️ ARCHIVED

**This repository is no longer actively maintained as a standalone package. The source code and future development have moved to the [htmltopdfwidgets monorepo](https://github.com/alihassan143/htmltopdfwidgets/tree/main/packages/docx_file_viewer).**

---

# Flutter DOCX Viewer Package

This Flutter package allows you to view DOCX files in your Flutter applications. It provides a simple way to load and display DOCX content in a Flutter app.

## Features

- Allows users to pick and view DOCX files.
- Displays the content of the DOCX file within your app.
- Tries to render DOCX content as accurately as possible, although some bugs may occur in certain files.

## Limitations

- The package may not render DOCX files perfectly in all cases.
- Some bugs are present in the rendering, and the file may not display as accurately as expected.
- The package strives to provide a functional rendering experience, but it may not be perfect for all DOCX files.

## Installation

To use this package, follow these steps:

1. Add the dependencies in your `pubspec.yaml` file:

    ```yaml
    dependencies:
      docx_file_viewer: ^0.0.1
    ```

2. Install the dependencies by running the following command:

    ```bash
    flutter pub get
    ```

## Example Usage

Below is an example of how to use the DOCX viewer in your Flutter application:

```dart
import 'dart:io';
import 'package:docx_file_viewer/docx_file_viewer.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'Docx Viewer'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  File? selectedFile;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
      ),
      body: selectedFile == null
          ? const Center(
              child: Text("Select File"),
            )
          : DocxViewer(
              file: selectedFile!,
            ),
      floatingActionButton: FloatingActionButton(
        onPressed: () async {
          final file = await FilePicker.platform.pickFiles();
          if (file == null) {
            return;
          }
          final filepath = file.files.first.path!;
          setState(() {
            selectedFile = File(filepath);
          });
        },
        tooltip: 'Select DOCX File',
        child: const Icon(Icons.add),
      ),
    );
  }
}
