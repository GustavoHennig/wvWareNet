# Porting Plan: Completing `decode_complex.c` Features in WvWareNet

## 1. Objective

This document outlines the tasks required to port the remaining text-centric parsing and formatting features from the original `wvWare/decode_complex.c` file into the `WvWareNet` C# project.

The primary goal is to accurately extract not just the text, but also the structural and formatting information that can be represented in a plain text format, such as lists, indentation, and basic text styling (bold, italic).

This plan intentionally **skips** features related to visual or graphical elements, such as images, shapes, page borders, and complex table shading.

## 2. Current Status

The C# `WordDocumentParser` has successfully implemented the foundational logic for parsing complex Word documents:

-   **Piece Table (CLX) Parsing**: The system can read the piece table to correctly sequence the document's text content.
-   **Basic Property Lookup**: The system reads the `PLCFBTEPAPX` (for paragraphs) and `PLCFBTECHPX` (for characters) to locate property information.
-   **Basic Text Extraction**: The main text content is extracted.
-   **Basic Styling**: A limited, hardcoded set of character properties (bold, italic) are being applied.

## 3. Detailed Porting Tasks

The following tasks are broken down into phases to incrementally build up the required functionality.

### Phase 1: Foundational - Comprehensive SPRM Handling

This phase focuses on creating a robust system for applying "sprms" (property modifications), which is the core mechanism Word uses for formatting.

-   [ ] **Task 1.1: Create Core SPRM Data Structures**
    -   [ ] Create a `Sprm.cs` class/record to represent a single property modification, storing its code, the target element (paragraph, character, section), and its value.
    -   [ ] Create a static `SprmDefinitions.cs` class to hold a mapping (e.g., a `Dictionary<ushort, SprmDefinition>`) from a numeric SPRM code to its definition (name, operand size, etc.). This will replace the `wvGetrgsprmPrm` and `wvGetrgsprmWord6` logic.

-   [ ] **Task 1.2: Develop the SPRM Applicator Service**
    -   [ ] Create an `ISprmApplicator.cs` interface and a `SprmApplicator.cs` implementation.
    -   [ ] Implement a method in `SprmApplicator` to parse a `grpprl` (a group of SPRMs) from a byte array, which is how they are stored in the piece table's `CLX` structure.
    -   [ ] This service will contain the logic to apply a list of parsed SPRMs to `Paragraph`, `Character`, and `Section` property objects.

-   [ ] **Task 1.3: Implement Paragraph Property (PAP) Application**
    -   [ ] Enhance the `ParagraphProperties.cs` model to include all text-relevant properties.
    -   [ ] In `SprmApplicator`, implement the logic for all `sgcPara` SPRMs.
    -   [ ] **Priority PAP SPRMs:**
        -   [ ] `sprmPJc`: Justification (left, right, center, justified).
        -   [ ] `sprmPInTable`, `sprmFTtp`: Flags indicating if the paragraph is part of a table.
        -   [ ] `sprmPDxaLeft`, `sprmPDxaRight`, `sprmPDxaFirstLine`: Indentation values (in twips).
        -   [ ] `sprmPDyaLine`: Line spacing.
        -   [ ] `sprmPIlfo`, `sprmPIlvl`: **Crucial for lists.** These identify the list and the level this paragraph belongs to.
        -   [ ] `sprmPIstd`: The base style index for the paragraph.
    -   [ ] **Skipped PAP SPRMs**: `sprmPShd` (shading), `sprmPBrd` (borders), `sprmPKeepLines` (widow/orphan control).

-   [ ] **Task 1.4: Implement Character Property (CHP) Application**
    -   [ ] Enhance the `CharacterProperties.cs` model to include all text-relevant properties.
    -   [ ] In `SprmApplicator`, implement the logic for all `sgcChp` SPRMs.
    -   [ ] **Priority CHP SPRMs:**
        -   [ ] `sprmCFBold`, `sprmCFItalic`, `sprmCFStrike`, `sprmCFSmallCaps`, `sprmCFVanish` (hidden), etc.
        -   [ ] `sprmCUnderline`: Underline style (single, double, dotted).
        -   [ ] `sprmCFontSize`: Font size.
        -   [ ] `sprmCIdslang`: Language ID for spell-checking/hyphenation.
    -   [ ] **Skipped CHP SPRMs**: `sprmCShd` (shading), `sprmCBord` (borders), `sprmCPic` (pictures).

### Phase 2: List and List Override Implementation

This phase builds on the SPRM foundation to correctly interpret and format numbered and bulleted lists.

-   [ ] **Task 2.1: Create List Data Models**
    -   [ ] Create `ListTable.cs` (`LST`) to store the definition of a single list, including its levels.
    -   [ ] Create `ListLevel.cs` (`LVL`) to define the properties of a single level within a list (e.g., number format, starting number, text template).
    -   [ ] Create `ListOverride.cs` (`LFO`) to handle overrides applied to a specific list instance.

-   [ ] **Task 2.2: Parse List Information from the Table Stream**
    -   [ ] In `WordDocumentParser`, implement `ParseListTables` to read the `PLCF LST` using the offsets from the `FIB`.
    -   [ ] Implement `ParseListOverrides` to read the `PLCF LFO` from the `FIB` offsets.

-   [ ] **Task 2.3: Develop a List Formatting Service**
    -   [ ] Create a `ListManager.cs` service.
    -   [ ] This service will hold the parsed list tables and overrides.
    -   [ ] It will track the current state of all active lists during document parsing (e.g., the current number for each level of each list).
    -   [ ] Implement a method `FormatParagraphPrefix(pap)` that takes a paragraph's properties, checks its `Ilfo` and `Ilvl` values, and returns the formatted string prefix (e.g., "1.", "a)", "  - ").

### Phase 3: Section and Document-Level Properties

-   [ ] **Task 3.1: Parse Section Properties (SEP)**
    -   [ ] Create a `SectionProperties.cs` model (`SEP`).
    -   [ ] In `WordDocumentParser`, parse the `PLCF SED` to get the boundaries of each section.
    -   [ ] Use the `SprmApplicator` to apply section-level SPRMs (`sgcSep`) to the `SectionProperties` object.
    -   [ ] **Priority SEP SPRMs**: `sprmSColumns` (number of columns).
    -   [ ] **Skipped SEP SPRMs**: Page margins, page size, headers/footers positioning.

-   [ ] **Task 3.2: Parse Document Properties (DOP)**
    -   [ ] Create a `DocumentProperties.cs` model (`DOP`).
    -   [ ] In `WordDocumentParser`, parse the `DOP` structure from the table stream.
    -   [ ] Use these properties as the global defaults for the document.

### Phase 4: Integration and Refactoring

-   [ ] **Task 4.1: Refactor `WordDocumentParser`**
    -   [ ] Modify the main parsing loop in `WordDocumentParser`.
    -   [ ] Instead of the current simple property handling, for each paragraph:
        1.  Assemble its properties by applying styles and SPRMs (using `SprmApplicator`).
        2.  If it's a list item (`pap.Ilfo > 0`), call the `ListManager` to get the list prefix.
        3.  Assemble the character runs with their respective properties.
        4.  Prepend the list prefix to the paragraph's text content.
        5.  Store the fully resolved properties in the `DocumentModel`.

-   [ ] **Task 4.2: Enhance Logging**
    -   [ ] Add detailed logging throughout the `SprmApplicator` and `ListManager` to trace how formatting is being resolved. Log the raw SPRMs being read and the resulting properties.

-   [ ] **Task 4.3: Final Verification**
    -   [ ] Test with a variety of documents containing complex lists, nested lists, and style overrides.
    -   [ ] Ensure the text output correctly represents the list structure and indentation.