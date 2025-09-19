import 'dart:io';

import 'package:docx_file_viewer/docx_file_viewer.dart';

class DocxViewer extends StatefulWidget {
  final File file;

  const DocxViewer({super.key, required this.file});

  @override
  State<DocxViewer> createState() => _DocxViewerState();
}

class _DocxViewerState extends State<DocxViewer> {
  late final Future<List<Widget>> loadDocument;
  late TransformationController _transformationController;
  bool _isZooming = false;

  @override
  void initState() {
    super.initState();
    _transformationController = TransformationController();
    loadDocument = DocxExtractor().renderLayout(widget.file);
  }

  @override
  void dispose() {
    _transformationController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: FutureBuilder(
        future: loadDocument,
        builder: (context, snapshot) {
          if (snapshot.hasData) {
            return GestureDetector(
              onScaleStart: (_) => setState(() => _isZooming = true), // Detect zoom gesture start
              onScaleEnd: (_) => setState(() => _isZooming = false),
              onScaleUpdate: (ScaleUpdateDetails details) {
                if (details.scale != 1.0) {
                  setState(() {
                    _isZooming = true; // Keep zooming enabled
                  });
                  _transformationController.value = Matrix4.identity()..scale(details.scale);
                }
              },
              child: InteractiveViewer(
                child: NotificationListener<ScrollNotification>(
                  onNotification: (ScrollNotification notification) {
                    return _isZooming; // Block scroll only when zooming
                  },
                  child: Scrollbar(
                    child: ListView.builder(
                      padding: const EdgeInsets.only(bottom: 68),
                      itemCount: snapshot.requireData.length,
                      physics: _isZooming ? const NeverScrollableScrollPhysics() : const BouncingScrollPhysics(),
                      itemBuilder: (_, int index) => snapshot.requireData[index],
                    ),
                  ),
                ),
              ),
            );
          } else if (snapshot.hasError) {
            return const Center(
              child: Text("Error rendering layout"),
            );
          } else {
            return const CircularProgressIndicator.adaptive();
          }
        },
      ),
    );
  }
}
