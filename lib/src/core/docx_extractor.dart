import 'dart:developer';
import 'dart:io';

import 'package:xml/xml.dart' as xml;

import '../../docx_file_viewer.dart';

class DocxExtractor {
  DocxExtractor._();
  static Future<List<Widget>> renderLayout(File file) async {
    try {
      final archive = ZipDecoder().decodeBytes(await file.readAsBytes());

      // Extract document.xml and document.xml.rels
      final documentXmlFile =
          archive.files.firstWhere((file) => file.name == 'word/document.xml');
      final relsXmlFile = archive.files
          .firstWhere((file) => file.name == 'word/_rels/document.xml.rels');

      // Parse XML
      final documentXml =
          xml.XmlDocument.parse(String.fromCharCodes(documentXmlFile.content));
      final relsXml =
          xml.XmlDocument.parse(String.fromCharCodes(relsXmlFile.content));
      // log(documentXml.toXmlString());
      // Extract image relationships
      final imageMap = _extractImageRelationships(relsXml, archive);

      // Parse the content
      return _parseContent(documentXml, imageMap);
    } catch (e) {
      log(e.toString());
      // Handle error, log it or provide a fallback widget

      return [
        const Text('Error parsing the document')
      ]; // Fallback widget in case of error
    }
  }

  static Map<String, Uint8List> _extractImageRelationships(
      xml.XmlDocument relsXml, Archive archive) {
    final imageMap = <String, Uint8List>{};

    relsXml.findAllElements('Relationship').forEach((rel) {
      final type = rel.getAttribute('Type') ?? '';
      final target = rel.getAttribute('Target') ?? '';
      final id = rel.getAttribute('Id') ?? '';

      if (type.contains('image')) {
        final filePath = 'word/$target';
        final file = archive.files.firstWhere(
          (file) => file.name == filePath,
        );
        imageMap[id] = Uint8List.fromList(file.content);
      }
    });

    return imageMap;
  }

  static Widget _parseParagraph(
      xml.XmlElement paragraph, Map<String, Uint8List> imageMap) {
    final spans = <InlineSpan>[];

    // Handle unordered or ordered list items
    final isListItem =
        paragraph.getElement('w:pPr')?.getElement('w:numPr') != null;
    final listLevel = _getListLevel(paragraph);

    if (isListItem) {
      // Handle list item with a bullet or number
      spans.add(TextSpan(
        text: _getListBullet(listLevel),
        style: const TextStyle(
            fontSize: 16, fontWeight: FontWeight.normal, color: Colors.black),
      ));
    }

    // Iterate through runs (text + style) in the paragraph
    paragraph.findAllElements('w:r').forEach((run) {
      log(run.innerText);
      final text = run.getElement('w:t')?.innerText ?? run.innerText;
      final innerText = run.innerText;
      final style = _parseRunStyle(run.getElement('w:rPr'));

      if (innerText.trim().isNotEmpty) {
        spans.add(TextSpan(
          text: innerText,
          style: style.copyWith(color: style.color ?? Colors.black),
        ));
      } else if (text.trim().isNotEmpty) {
        spans.add(TextSpan(
          text: text,
          style: style.copyWith(color: style.color ?? Colors.black),
        ));
      }

      // Check for embedded images
      run.findAllElements('a:blip').forEach((imageElement) {
        final embedId = imageElement.getAttribute('r:embed') ?? '';
        if (imageMap.containsKey(embedId)) {
          spans.add(WidgetSpan(
            child: Padding(
              padding: const EdgeInsets.symmetric(vertical: 8.0),
              child: Image.memory(imageMap[embedId]!),
            ),
          ));
        }
      });
    });

    // Handle headings
    final headingStyle = _parseHeadingStyle(paragraph);
    if (headingStyle != null) {
      return Padding(
        padding: const EdgeInsets.only(bottom: 12.0),
        child: RichText(
          text: TextSpan(
            children: spans,
            style: headingStyle,
          ),
        ),
      );
    }

    // Handle page breaks
    final hasPageBreak =
        paragraph.findElements('w:lastRenderedPageBreak').isNotEmpty;
    if (hasPageBreak) {
      return const SizedBox(height: 20); // Provide spacing for a page break
    }

    // Handle paragraph spacing
    final paragraphSpacing = _parseParagraphSpacing(paragraph);

    return Padding(
      padding: paragraphSpacing,
      child: RichText(
        text: TextSpan(
          children: spans,
          style: const TextStyle(fontSize: 16),
        ),
      ),
    );
  }

  static EdgeInsets _parseParagraphSpacing(xml.XmlElement paragraph) {
    final pPr = paragraph.getElement('w:pPr');
    final before = int.tryParse(
            pPr?.getElement('w:spacing')?.getAttribute('w:before') ?? "0") ??
        0;
    final after = int.tryParse(
            pPr?.getElement('w:spacing')?.getAttribute('w:after') ?? "0") ??
        0;

    // Convert Word spacing units to Flutter padding (assume 20 units = 1 point)
    return EdgeInsets.only(
      top: before / 20,
      bottom: after / 20,
    );
  }

  static TextStyle? _parseHeadingStyle(xml.XmlElement paragraph) {
    final pStyle = paragraph
        .getElement('w:pPr')
        ?.getElement('w:pStyle')
        ?.getAttribute('w:val');
    if (pStyle != null) {
      switch (pStyle) {
        case 'Heading1':
          return const TextStyle(fontSize: 32, fontWeight: FontWeight.bold);
        case 'Heading2':
          return const TextStyle(fontSize: 28, fontWeight: FontWeight.bold);
        case 'Heading3':
          return const TextStyle(fontSize: 24, fontWeight: FontWeight.bold);
        case 'Heading4':
          return const TextStyle(fontSize: 20, fontWeight: FontWeight.bold);
        case 'Heading5':
          return const TextStyle(fontSize: 18, fontWeight: FontWeight.bold);
        case 'Heading6':
          return const TextStyle(fontSize: 16, fontWeight: FontWeight.bold);
        default:
          return null; // Default to body text if not a heading
      }
    }
    return null; // Not a heading
  }

  static Widget _parseUnorderedList(xml.XmlElement list) {
    final listItems = <Widget>[];

    list.findAllElements('w:p').forEach((item) {
      listItems.add(Padding(
        padding: const EdgeInsets.only(left: 20.0),
        child: Row(
          children: [
            const Icon(Icons.circle, size: 6), // Bullet icon
            const SizedBox(width: 8),
            Expanded(child: Text(item.innerText)),
          ],
        ),
      ));
    });

    return Column(children: listItems);
  }

  static String _getListBullet(int level) {
    // Customize list bullet or numbering based on the level
    // Simple bullet for level 0, adjust for numbering as needed
    if (level == 0) {
      return '• '; // Bullet for level 0
    }
    // For other levels, return numbered list (you can customize this)
    return '${level + 1}. ';
  }

  static int _getListLevel(xml.XmlElement paragraph) {
    // Extract the list level from the paragraph
    final numPr = paragraph.getElement('w:pPr')?.getElement('w:numPr');
    final ilvl = num.tryParse(
            numPr?.getElement('w:ilvl')?.getAttribute('w:val') ?? "0") ??
        0;
    return ilvl.toInt();
  }

  static Widget _parseOrderedList(xml.XmlElement list) {
    final listItems = <Widget>[];
    int counter = 1;

    list.findAllElements('w:p').forEach((item) {
      listItems.add(Padding(
        padding: const EdgeInsets.only(left: 20.0),
        child: Row(
          children: [
            Text('$counter. '),
            Expanded(child: Text(item.innerText)),
          ],
        ),
      ));
      counter++;
    });

    return Column(children: listItems);
  }

  static List<Widget> _parseContent(
      xml.XmlDocument documentXml, Map<String, Uint8List> imageMap) {
    final widgets = <Widget>[];

    for (final body in documentXml.findAllElements('w:body')) {
      for (final element in body.children.whereType<xml.XmlElement>()) {
        log(element.name.local);
        switch (element.name.local) {
          case 'p':
            widgets.add(_parseParagraph(element, imageMap));
            break;
          case 'tbl':
            widgets.add(_parseTable(element, imageMap));
            break;
          case 'ul':
            widgets.add(_parseUnorderedList(element));
            break;
          case 'ol':
            widgets.add(_parseOrderedList(element));
            break;
          case 'sdt':
            widgets.add(_parseSdt(element, imageMap));
            break;
          // case 'sectPr':
          //   widgets.add(_parseSectionProperties(element));
          // break;
        }
      }
    }

    return widgets;
  }

  static Widget _parseTable(
      xml.XmlElement table, Map<String, Uint8List> imageMap) {
    final rows = <TableRow>[];

    // Parse table border properties
    final borderStyle = _parseTableBorderStyle(table);

    // Store the number of cells in the first row to ensure all rows match this count
    int maxCells = 0;

    table.findAllElements('w:tr').forEach((row) {
      final cells = <Widget>[];

      // Parse cells in each row
      row.findAllElements('w:tc').forEach((cell) {
        final cellContent = <Widget>[];

        // Check for background fill color (shading)
        final backgroundColor = _parseCellBackgroundColor(cell);

        cell.findAllElements('w:p').forEach((paragraph) {
          cellContent.add(_parseParagraph(paragraph, imageMap));
        });

        cells.add(Container(
          padding: const EdgeInsets.all(8.0),
          decoration: BoxDecoration(
            color: backgroundColor, // Apply background color
            border:
                Border.all(color: borderStyle.color, width: borderStyle.width),
          ),
          child: Column(children: cellContent),
        ));
      });

      // Update the maxCells if this row has more cells
      maxCells = maxCells > cells.length ? maxCells : cells.length;

      // Add the row to the table
      rows.add(TableRow(children: cells));
    });

    // Add empty cells if a row has fewer cells than maxCells
    for (var row in rows) {
      final childrenCount = row.children.length;
      if (childrenCount < maxCells) {
        for (int i = 0; i < maxCells - childrenCount; i++) {
          row.children.add(Container(
            padding: const EdgeInsets.all(8.0),
            child: const SizedBox.shrink(), // Empty cell
          ));
        }
      }
    }

    return Table(
      border:
          TableBorder.all(color: borderStyle.color, width: borderStyle.width),
      children: rows,
    );
  }

  // Parse table border style (color, width)
  static TableBorderStyle _parseTableBorderStyle(xml.XmlElement table) {
    final borderColor = table
            .getElement('w:tblBorders')
            ?.getElement('w:top')
            ?.getAttribute('w:color') ??
        '000000';
    final borderWidth = table
            .getElement('w:tblBorders')
            ?.getElement('w:top')
            ?.getAttribute('w:space') ??
        '1';

    return TableBorderStyle(
      color: _hexToColor(borderColor),
      width: double.parse(borderWidth),
    );
  }

  // Parse background color (shading) for a table cell
  static Color _parseCellBackgroundColor(xml.XmlElement cell) {
    final shading = cell.getElement('w:shd');
    final fillColor = shading?.getAttribute('w:fill');

    if (fillColor != null) {
      return _hexToColor(fillColor);
    } else {
      return Colors.transparent; // No background color
    }
  }

  static Widget _parseSdt(
      xml.XmlElement sdtElement, Map<String, Uint8List> imageMap) {
    final content = sdtElement
        .findAllElements('w:sdtContent')
        .expand((contentElement) => contentElement.children)
        .whereType<xml.XmlElement>();

    final contentWidgets = content.map((childElement) {
      switch (childElement.name.local) {
        case 'p':
          return _parseParagraph(childElement, imageMap);
        case 'tbl':
          return _parseTable(childElement, imageMap);
        default:
          return Text(childElement.innerText);
      }
    }).toList();

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: contentWidgets,
    );
  }

  // static Widget _parseSectionProperties(xml.XmlElement sectPrElement) {
  //   final pageSettings = <Widget>[];

  //   // Extract page size
  //   final pgSz = sectPrElement.getElement('w:pgSz');
  //   if (pgSz != null) {
  //     final width = pgSz.getAttribute('w:w');
  //     final height = pgSz.getAttribute('w:h');
  //     if (width != null && height != null) {
  //       pageSettings.add(Text('Page Size: ${width}x$height twips'));
  //     }
  //   }

  //   // Extract margins
  //   final pgMar = sectPrElement.getElement('w:pgMar');
  //   if (pgMar != null) {
  //     final top = pgMar.getAttribute('w:top');
  //     final bottom = pgMar.getAttribute('w:bottom');
  //     final left = pgMar.getAttribute('w:left');
  //     final right = pgMar.getAttribute('w:right');
  //     pageSettings.add(Text(
  //         'Margins - Top: $top, Bottom: $bottom, Left: $left, Right: $right'));
  //   }

  //   return Column(
  //     crossAxisAlignment: CrossAxisAlignment.start,
  //     children: pageSettings,
  //   );
  // }

  static TextStyle _parseRunStyle(xml.XmlElement? styleElement) {
    if (styleElement == null) return const TextStyle();

    // Check for basic style properties (bold, italic, underline)
    final isBold = styleElement.findElements('w:b').isNotEmpty;
    final isItalic = styleElement.findElements('w:i').isNotEmpty;
    final isUnderline = styleElement.findElements('w:u').isNotEmpty;

    // Parse font size (half-points to points)
    final fontSize = double.tryParse(
          styleElement.getElement('w:sz')?.getAttribute('w:val') ?? '32',
        ) ??
        16.0;
    // Parse text color (if available)
    final colorHex =
        styleElement.findElements('w:color').firstOrNull?.getAttribute('w:val');
    final textColor = colorHex != null ? _hexToColor(colorHex) : Colors.black;

    // Ensure transparent colors are assigned black
    final effectiveTextColor =
        textColor == Colors.transparent ? Colors.black : textColor;

    // Parse background color (shading) for text background
    final shadingElement = styleElement.findElements('w:shd').firstOrNull;
    final backgroundColor = shadingElement?.getAttribute('w:fill') != null
        ? _hexToColor(shadingElement!.getAttribute('w:fill')!)
        : Colors.transparent;

    return TextStyle(
      fontWeight: isBold ? FontWeight.bold : FontWeight.normal,
      fontStyle: isItalic ? FontStyle.italic : FontStyle.normal,
      decoration: isUnderline ? TextDecoration.underline : TextDecoration.none,
      fontSize: fontSize / 2, // Word font size is in half-points
      color: effectiveTextColor, // Set the effective text color
      backgroundColor: backgroundColor, // Set the background color
    );
  }

  static Color _hexToColor(String hex) {
    // Clean up the hex string by removing any '#' and converting it to uppercase
    final hexColor = hex.replaceAll('#', '').toUpperCase();

    // If the hex code is 6 characters long (RGB), make it 8 by adding full opacity (FF)
    if (hexColor.length == 6) {
      return Color(int.parse('0xFF$hexColor'));
    }

    // If the hex code is 8 characters long (RGBA), use it directly
    if (hexColor.length == 8) {
      // Extract the alpha value (first 2 characters)
      final alpha = int.parse(hexColor.substring(0, 2), radix: 16);

      // If the alpha value is 0 (fully transparent), return black color
      if (alpha == 0) {
        return Colors.black; // Fallback to black
      }

      return Color(int.parse('0x$hexColor'));
    }

    // Fallback for unexpected formats (return black in case of an invalid format)
    return Colors.black;
  }
}

class TableBorderStyle {
  final Color color;
  final double width;

  TableBorderStyle({required this.color, required this.width});
}
