from checks.formatting.bibliography_style_check import BibliographyStyleCheck
from checks.formatting.caption_style_check import CaptionStyleCheck
from checks.formatting.content_style_check import ContentHeadingStyleCheck
from checks.formatting.cover_styles_check import CoverStylesCheck
from checks.formatting.custom_style_inheritance_check import CustomStyleInheritanceCheck
from checks.formatting.custom_style_usage_check import RequiredCustomStylesUsageCheck
from checks.formatting.custom_style_with_tabs_check import CustomStyleWithTabsCheck
from checks.formatting.frontpage_styles_check import FrontpageStylesCheck
from checks.formatting.heading_hierarchical_numbering_check import HeadingHierarchicalNumberingCheck
from checks.formatting.heading_style_check import HeadingStyleCheck
from checks.formatting.headings_used_corretcly_check import HeadingsUsedCorrectlyCheck
from checks.formatting.list_level_2_used_check import ListLevel2UsedCheck
from checks.formatting.main_chapter_starts_on_new_page_check import MainChapterStartsOnNewPageCheck
from checks.formatting.manual_horizontal_formatting_check import ManualHorizontalSpacingCheck
from checks.formatting.manual_vertical_formatting_check import ManualVerticalSpacingCheck
from checks.formatting.original_formatting_check import OriginalFormattingCheck
from checks.formatting.toc_heading_numbering_check import TocHeadingNumberingCheck
from checks.formatting.unnumbered_special_headings_check import UnnumberedSpecialHeadingsCheck
from checks.objects.image_low_quality_check import ImageLowQualityCheck
from checks.objects.list_of_figures_not_up_to_date_check import ListOfFiguresNotUpdatedCheck
from checks.objects.missing_list_of_fugures_check import MissingListOfFiguresCheck
from checks.objects.object_caption_bindings_check import ObjectCaptionBindingCheck
from checks.objects.object_caption_check import ObjectCaptionCheck
from checks.objects.object_caption_description_check import ObjectCaptionDescriptionCheck
from checks.objects.object_cross_reference_check import ObjectCrossReferenceCheck
from checks.structure.chapter_numbering_continuity_check import ChapterNumberingContinuityCheck
from checks.structure.document_structure_check import DocumentStructureCheck
from checks.structure.first_chapter_page1_check import FirstChapterStartsOnPageOneCheck
from checks.structure.toc_exists_check import TOCExistsCheck
from checks.structure.toc_first_section_check import TOCFirstSectionContentCheck
from checks.structure.toc_heading_levels_check import TOCHeadingLevelsCheck
from checks.structure.toc_illegal_content_check import TOCIllegalContentCheck
from checks.structure.toc_up_to_date_check import TOCUpToDateCheck
from document.word_document import WordDocument
from core.runner import Runner
from core.report import Report

# --- section checks ---
from checks.sections.section_count_check import SectionCountCheck
from checks.sections.section1_toc_check import Section1TOCCheck
from checks.sections.section2_text_check import Section2TextCheck
from checks.sections.section3_figure_list_check import Section3FigureListCheck
from checks.sections.section3_table_list_check import Section3TableListCheck
from checks.sections.section3_bibliography_check import Section3BibliographyCheck

# --- assignment ---
from assignment.assignment_loader import load_assignment

# --- formatting (assignment-based) ---
from checks.formatting.normal_style_check import NormalStyleCheck


def main():
    # 1) Načtení dokumentu a zadání
    doc = WordDocument("student.docx")
    assignment = load_assignment("assignment/assignment.json")
    doc.save_xml()

    # 2) Seznam kontrol (pořadí odpovídá tabulce hodnocení)
    checks = [
        # --- Části / oddíly dokumentu ---
        # SectionCountCheck(),
        # Section1TOCCheck(),
        # Section2TextCheck(),
        # Section3FigureListCheck(),
        # Section3TableListCheck(),
        # Section3BibliographyCheck(),

        # --- Formátování dle zadání ---
        # NormalStyleCheck(),
        # HeadingStyleCheck(1),
        # HeadingStyleCheck(2),
        # HeadingStyleCheck(3),
        # HeadingHierarchicalNumberingCheck(),
        # TocHeadingNumberingCheck()
        # UnnumberedSpecialHeadingsCheck(),
        # CoverStylesCheck(),
        # FrontpageStylesCheck(),
        # BibliographyStyleCheck(),
        # CaptionStyleCheck(),
        # ContentHeadingStyleCheck()
        #HeadingsUsedCorrectlyCheck()
        #OriginalFormattingCheck()
        # CustomStyleInheritanceCheck(),
        # RequiredCustomStylesUsageCheck(),
        # CustomStyleWithTabsCheck(),
        # MainChapterStartsOnNewPageCheck(),
        # ManualHorizontalSpacingCheck(),
        # ManualVerticalSpacingCheck(),
        # ListLevel2UsedCheck(),
        # InconsistentFormattingCheck()
        # -------Obsah a struktura-----
        # TOCExistsCheck(),
        # TOCUpToDateCheck(),
        # DocumentStructureCheck(),
        # TOCHeadingLevelsCheck(),
        # TOCFirstSectionContentCheck(),
        # TOCIllegalContentCheck(),
        # FirstChapterStartsOnPageOneCheck(),
        # ChapterNumberingContinuityCheck(),
        #--------Objekty-----------------
        # MissingListOfFiguresCheck(),
        # ListOfFiguresNotUpdatedCheck(),
        # ImageLowQualityCheck(),
        # ObjectCaptionCheck(),
        # ObjectCaptionDescriptionCheck(),
        # ObjectCrossReferenceCheck(),
        ObjectCaptionBindingCheck()
    ]

    # 3) Spuštění kontrol
    runner = Runner(checks)
    report = Report()

    results = runner.run(doc, assignment)

    # 4) Výpis výsledků
    for check, result in results:
        report.add(check.name, result)

    report.print()

    # 5) Debug (volitelné)
    # print("\nFIELD INSTRUCTIONS IN SECTION 3:")
    # print(doc.get_field_instructions(doc.section(2)))


if __name__ == "__main__":
    main()