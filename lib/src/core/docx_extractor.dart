import 'dart:convert';
import 'dart:developer';
import 'dart:io';

import 'package:xml/xml.dart';

import '../../docx_file_viewer.dart';

class DocxExtractor {
  Future<List<Widget>> renderLayout(File file) async {
    try {
      Map<String, Map<int, String>> numberingMap = {};
      final archive = ZipDecoder().decodeBytes(await file.readAsBytes());

      // Extract document.xml and document.xml.rels
      final documentXmlFile =
          archive.files.firstWhere((file) => file.name == 'word/document.xml');
      final numberingXmlFile = archive.files
          .where((file) => file.name == 'word/numbering.xml')
          .firstOrNull;

      final relsXmlFile = archive.files
          .firstWhere((file) => file.name == 'word/_rels/document.xml.rels');
      // final stylesXmlFile =
      //     archive.files.firstWhere((file) => file.name == 'word/styles.xml');
      // Parse XML
      final documentXml =
          XmlDocument.parse(String.fromCharCodes(documentXmlFile.content));
      final relsXml =
          XmlDocument.parse(String.fromCharCodes(relsXmlFile.content));
      if (numberingXmlFile != null) {
        final numberingXmlContent =
            utf8.decode(numberingXmlFile.content as List<int>);
        numberingMap = parseNumberingDefinitions(numberingXmlContent);
      }
      final imageMap = _extractImageRelationships(relsXml, archive);
      // await _loadFontsFromFontTable(archive);
      // await _loadStyles(stylesXml);
      // Parse the content
      return _parseContent(documentXml, imageMap, numberingMap);
    } catch (e) {
      log(e.toString());
      // Handle error, log it or provide a fallback widget

      return [
        const Text('Error parsing the document')
      ]; // Fallback widget in case of error
    }
  }

  Map<String, Map<int, String>> parseNumberingDefinitions(String numberingXml) {
    final document = XmlDocument.parse(numberingXml);
    final numberingMap = <String, Map<int, String>>{};

    // Iterate through all abstractNum elements
    for (final abstractNum in document.findAllElements('w:abstractNum')) {
      final abstractNumId = abstractNum.getAttribute('w:abstractNumId') ?? '';
      final levels = <int, String>{};

      for (final level in abstractNum.findAllElements('w:lvl')) {
        final ilvl = int.tryParse(level.getAttribute('w:ilvl') ?? '0') ?? 0;
        final format =
            level.getElement('w:numFmt')?.getAttribute('w:val') ?? '';

        levels[ilvl] = format;
      }

      numberingMap[abstractNumId] = levels;
    }

    return numberingMap; // Map of abstractNumId to level definitions
  }

  Map<String, Uint8List> _extractImageRelationships(
      XmlDocument relsXml, Archive archive) {
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

  Widget _parseParagraph(
      XmlElement paragraph,
      Map<String, Uint8List> imageMap,
      Map<String, Map<int, String>> numberingDefinitions,
      Map<String, int> counter) {
    final spans = <InlineSpan>[];

    // Handle unordered or ordered list items
    final isListItem =
        paragraph.getElement('w:pPr')?.getElement('w:numPr') != null;

    if (isListItem) {
      final numPr = paragraph.getElement('w:pPr')?.getElement('w:numPr');
      final numId = numPr?.getElement('w:numId')?.getAttribute('w:val');
      final ilvl = int.tryParse(
            numPr?.getElement('w:ilvl')?.getAttribute('w:val') ?? '0',
          ) ??
          0;

      if (numId != null && numberingDefinitions.containsKey(numId)) {
        final listLevelStyle = numberingDefinitions[numId]?[ilvl];
        int newLevel = level ?? ilvl;

        if (listLevelStyle != null) {
          if (listLevelStyle.toLowerCase() == 'bullet') {
            level = null;
          }
          spans.add(WidgetSpan(
            child: Padding(
              padding: EdgeInsets.only(left: ilvl != 0 ? 30 : 20.0),
              child: listLevelStyle.toLowerCase() == 'bullet'
                  ? Text(
                      _getBulletForLevel(ilvl),
                      style: const TextStyle(
                        fontSize: 20, // Adjust size based on level
                        color: Colors.black,
                      ),
                    )
                  : Text(
                      _getFormattedListNumber(
                          numberingDefinitions, counter, numId, newLevel),
                      style: const TextStyle(color: Colors.black, fontSize: 16),
                    ),
            ),
          ));
        }
      }
    } else {
      level = null;
      enabled = false;
    }

    // Iterate through runs (text + style) in the paragraph
    paragraph.findAllElements('w:r').forEach((run) {
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
      final hasPageBreak = run.getElement('w:pict');

      if (hasPageBreak != null) {
        spans.add(const WidgetSpan(
            child: SizedBox(
          width: double.infinity,
          child: Divider(
            color: Colors.grey,
            thickness: 1,
          ),
        )));
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
    // Handle paragraph spacing
    final paragraphSpacing = _parseParagraphSpacing(paragraph);
    final textAlignment =
        getTextAlign(paragraph.getElement('w:pPr')?.getElement('w:jc'));

    if (headingStyle != null) {
      return Padding(
        padding: paragraphSpacing,
        child: RichText(
          textAlign: textAlignment,
          text: TextSpan(
            children: spans,
            style: headingStyle,
          ),
        ),
      );
    }

    // Handle paragraph spacing

    return Padding(
      padding: paragraphSpacing,
      child: RichText(
        textAlign: textAlignment,
        text: TextSpan(
          children: spans,
          style: const TextStyle(fontSize: 16),
        ),
      ),
    );
  }

  TextAlign getTextAlign(XmlElement? alignment) {
    final String? alignementValue = alignment?.getAttribute("w:val");
    log(alignementValue.toString());
    if (alignementValue == null) {
      return TextAlign.start;
    }
    switch (alignementValue) {
      case "left":
        return TextAlign.left;
      case "center":
        return TextAlign.center;
      case "start":
        return TextAlign.start;
      case "end":
        return TextAlign.end;
      case "right":
        return TextAlign.right;
      case "both":
        return TextAlign.justify;
      default:
        return TextAlign.start;
    }
  }

  EdgeInsets _parseParagraphSpacing(XmlElement paragraph) {
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

  TextStyle? _parseHeadingStyle(XmlElement paragraph) {
    final pStyle = paragraph
        .getElement('w:pPr')
        ?.getElement('w:pStyle')
        ?.getAttribute('w:val');
    TextStyle style = _parseRunStyle(paragraph.getElement('w:rPr'));

    if (pStyle != null) {
      switch (pStyle) {
        case 'Heading1':
          style = style.copyWith(fontSize: 32, fontWeight: FontWeight.bold);
        case 'Heading2':
          style = style.copyWith(fontSize: 28, fontWeight: FontWeight.bold);
        case 'Heading3':
          style = style.copyWith(fontSize: 24, fontWeight: FontWeight.bold);
        case 'Heading4':
          style = style.copyWith(fontSize: 20, fontWeight: FontWeight.bold);
        case 'Heading5':
          style = style.copyWith(fontSize: 18, fontWeight: FontWeight.bold);
        case 'Heading6':
          style = style.copyWith(fontSize: 16, fontWeight: FontWeight.bold);
        default:
          break;
      }
    }

    return style; // Not a heading
  }

// Helper to determine the list type

// Helper to get list bullet (number or bullet symbol)
  String _getBulletForLevel(int level) {
    const bulletStyles = ['•', '◦', '▪', '▫', '»', '›', '⁃', '–'];
    return bulletStyles[level % bulletStyles.length];
  }

  int? level;
  bool enabled = false;

  /// Retrieves the formatted list number for an ordered list item
  String _getFormattedListNumber(
      Map<String, Map<int, String>> numberingDefinitions,
      Map<String, int> counters,
      String numId,
      int ilvl) {
    // Retrieve the format for the current level
    final format = numberingDefinitions[numId]?[ilvl] ?? 'decimal';

    // Increment the counter for the given numId and ilvl
    final key = '$numId-$ilvl';
    counters[key] = (counters[key] ?? 0) + 1;
    final number = counters[key]!;

    // Generate list number based on format
    switch (format.toLowerCase()) {
      case 'decimal':
        return '$number.'; // Example: 1., 2., 3.
      case 'lowerroman':
        return '${_toRoman(number).toLowerCase()}.'; // Example: i., ii., iii.
      case 'upperroman':
        return '${_toRoman(number).toUpperCase()}.'; // Example: I., II., III.
      case 'lowerletter':
        return '${String.fromCharCode(97 + (number - 1))}.'; // Example: a., b., c.
      case 'upperletter':
        return '${String.fromCharCode(65 + (number - 1))}.'; // Example: A., B., C.
      default:
        return '$number.'; // Fallback to decimal
    }
  }

  String _toRoman(int number) {
    final romanNumerals = [
      'M',
      'CM',
      'D',
      'CD',
      'C',
      'XC',
      'L',
      'XL',
      'X',
      'IX',
      'V',
      'IV',
      'I'
    ];
    final romanValues = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
    var result = '';
    var i = 0;

    while (number > 0) {
      while (number >= romanValues[i]) {
        result += romanNumerals[i];
        number -= romanValues[i];
      }
      i++;
    }
    return result;
  }

  /// Retrieves the list type (ordered/unordered) based on paragraph properties

  List<Widget> _parseContent(
      XmlDocument documentXml,
      Map<String, Uint8List> imageMap,
      Map<String, Map<int, String>> numberingDefinitions) {
    final widgets = <Widget>[];
    final counters = <String, int>{};
    for (final body in documentXml.findAllElements('w:body')) {
      for (final element in body.children.whereType<XmlElement>()) {
        // log(element.name.local);
        switch (element.name.local) {
          case 'p':
            widgets.add(_parseParagraph(
                element, imageMap, numberingDefinitions, counters));
            break;
          case 'tbl':
            widgets.add(
                _parseTable(element, imageMap, numberingDefinitions, counters));
            break;

          case 'sdt':
            widgets.add(
                _parseSdt(element, imageMap, numberingDefinitions, counters));
            break;
          // case 'sectPr':
          //   widgets.add(_parseSectionProperties(element));
          // break;
        }
      }
    }

    return widgets;
  }

  Widget _parseTable(
      XmlElement table,
      Map<String, Uint8List> imageMap,
      Map<String, Map<int, String>> numberingDefinitions,
      Map<String, int> counter) {
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
          cellContent.add(_parseParagraph(
              paragraph, imageMap, numberingDefinitions, counter));
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
  TableBorderStyle _parseTableBorderStyle(XmlElement table) {
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
  Color _parseCellBackgroundColor(XmlElement cell) {
    final shading = cell.getElement('w:shd');
    final fillColor = shading?.getAttribute('w:fill');

    if (fillColor != null) {
      return _hexToColor(fillColor);
    } else {
      return Colors.transparent; // No background color
    }
  }

  Widget _parseSdt(
      XmlElement sdtElement,
      Map<String, Uint8List> imageMap,
      Map<String, Map<int, String>> numberingDefinitions,
      Map<String, int> counter) {
    final content = sdtElement
        .findAllElements('w:sdtContent')
        .expand((contentElement) => contentElement.children)
        .whereType<XmlElement>();

    final contentWidgets = content.map((childElement) {
      switch (childElement.name.local) {
        case 'p':
          return _parseParagraph(
              childElement, imageMap, numberingDefinitions, counter);
        case 'tbl':
          return _parseTable(
              childElement, imageMap, numberingDefinitions, counter);
        default:
          return Text(childElement.innerText);
      }
    }).toList();

    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: contentWidgets,
    );
  }

  TextStyle _parseRunStyle(XmlElement? styleElement) {
    if (styleElement == null) return const TextStyle();

    final isBold = styleElement.findElements('w:b').isNotEmpty;
    final isItalic = styleElement.findElements('w:i').isNotEmpty;
    final isUnderline = styleElement.findElements('w:u').isNotEmpty;

    // Parse font size (half-points to points)
    final fontSize = double.tryParse(
          styleElement.getElement('w:sz')?.getAttribute('w:val') ?? '20',
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

    String? bgHex;
    final highlightElement = styleElement.getElement('w:highlight');
    final shadingElement = styleElement.getElement('w:shd');

    if (highlightElement != null) {
      bgHex = highlightElement.getAttribute('w:val');
    } else if (shadingElement != null) {
      bgHex = shadingElement.getAttribute('w:fill');
    }
    final backgroundColor = bgHex != null && bgHex != 'auto'
        ? _hexToColor(bgHex)
        : Colors.transparent;

    // If custom font is available, load it
    // final fontId = styleElement.getElement('w:rFonts')?.getAttribute('w:ascii');

    // final fontFamily = _fontNameMapping[fontId ?? ''] ?? 'Roboto';
    // log(fontFamily);
    return TextStyle(
      // fontFamily: fontFamily,
      fontWeight: isBold ? FontWeight.bold : FontWeight.normal,
      fontStyle: isItalic ? FontStyle.italic : FontStyle.normal,
      decoration: isUnderline ? TextDecoration.underline : TextDecoration.none,
      fontSize: fontSize / 2, // Word font size is in half-points
      color: effectiveTextColor, // Set the effective text color
      backgroundColor: backgroundColor, // Set the background color
    );
  }

  Color _hexToColor(String hex) {
    try {
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
    } catch (e) {
      log(hex);
      return getColorFromString(hex);
    }
  }

  Color getColorFromString(String colorName) {
    // Define a map of supported colors
    Map<String, Color> colorMap = {
      'yellow': Colors.yellow,
      'red': Colors.red,
      'blue': Colors.blue,
      'green': Colors.green,
      'black': Colors.black,
      'white': Colors.white,
      'orange': Colors.orange,
      'purple': Colors.purple,
      'pink': Colors.pink,
      'brown': Colors.brown,
      'cyan': Colors.cyan,
      'grey': Colors.grey,
    };

    // Return the color if found, else default to transparent
    return colorMap[colorName.toLowerCase()] ?? Colors.transparent;
  }
}

class TableBorderStyle {
  final Color color;
  final double width;

  TableBorderStyle({required this.color, required this.width});
}
