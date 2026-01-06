from assignment.excel.excel_assignment_loader import load_excel_assignment
from checks.excel.chart.missing_chart_check import MissingChartCheck
from checks.excel.data_process.array_formula_check import ArrayFormulaCheck
from checks.excel.data_process.descriptive_statistic_check import DescriptiveStatisticsCheck
from checks.excel.data_process.missing_wrong_formula_check import MissingOrWrongFormulaOrNotCalculatedCheck
from checks.excel.data_process.named_range_usage_check import NamedRangeUsageCheck
from checks.excel.data_process.redundant_absolute_reference_check import RedundantAbsoluteReferenceCheck
from checks.excel.data_process.required_data_worksheet_check import RequiredDataWorksheetCheck
from checks.excel.data_process.required_source_worksheet_check import RequiredSourceWorksheetCheck
from checks.excel.data_process.non_copyable_formula_check import NonCopyableFormulasCheck
from checks.excel.formatting.cells_merge_check import MergedCellsCheck
from checks.excel.formatting.conditional_formatting_check import ConditionalFormattingExistsCheck
from checks.excel.formatting.conditional_formatting_is_correct_check import ConditionalFormattingCorrectnessCheck
from checks.excel.formatting.header_formatting_check import HeaderFormattingCheck
from checks.excel.formatting.number_formatting_check import NumberFormattingCheck
from checks.excel.formatting.table_border_check import TableBorderCheck
from checks.word.bibliography.bibliography_exist_check import MissingBibliographyCheck
from checks.word.bibliography.bibliography_up_to_date_check import BibliographyNotUpdatedCheck
from checks.word.formatting.bibliography_style_check import BibliographyStyleCheck
from checks.word.formatting.caption_style_check import CaptionStyleCheck
from checks.word.formatting.content_style_check import ContentHeadingStyleCheck
from checks.word.formatting.cover_styles_check import CoverStylesCheck
from checks.word.formatting.custom_style_inheritance_check import CustomStyleInheritanceCheck
from checks.word.formatting.custom_style_usage_check import RequiredCustomStylesUsageCheck
from checks.word.formatting.custom_style_with_tabs_check import CustomStyleWithTabsCheck
from checks.word.formatting.frontpage_styles_check import FrontpageStylesCheck
from checks.word.formatting.heading_hierarchical_numbering_check import HeadingHierarchicalNumberingCheck
from checks.word.formatting.heading_style_check import HeadingStyleCheck
from checks.word.formatting.headings_used_corretcly_check import HeadingsUsedCorrectlyCheck
from checks.word.formatting.incosistent_formatting_check import InconsistentFormattingCheck
from checks.word.formatting.list_level_2_used_check import ListLevel2UsedCheck
from checks.word.formatting.main_chapter_starts_on_new_page_check import MainChapterStartsOnNewPageCheck
from checks.word.formatting.manual_horizontal_formatting_check import ManualHorizontalSpacingCheck
from checks.word.formatting.manual_vertical_formatting_check import ManualVerticalSpacingCheck
from checks.word.formatting.original_formatting_check import OriginalFormattingCheck
from checks.word.formatting.toc_heading_numbering_check import TocHeadingNumberingCheck
from checks.word.formatting.unnumbered_special_headings_check import UnnumberedSpecialHeadingsCheck
from checks.word.header_footer.section_emty_footer_check import SectionFooterEmptyCheck
from checks.word.header_footer.section_emty_header_check import SectionHeaderEmptyCheck
from checks.word.header_footer.header_footer_missing_check import HeaderFooterMissingCheck
from checks.word.header_footer.section_footer_linked_check import FooterLinkedToPreviousCheck
from checks.word.header_footer.section_footer_page_number_check import SectionFooterHasPageNumberCheck
from checks.word.header_footer.section_header_linked_check import HeaderNotLinkedToPreviousCheck
from checks.word.header_footer.second_section_header_text_check import SecondSectionHeaderHasTextCheck
from checks.word.header_footer.second_section_page_num_start_at_one_check import SecondSectionPageNumberStartsAtOneCheck
from checks.word.objects.image_low_quality_check import ImageLowQualityCheck
from checks.word.objects.list_of_figures_not_up_to_date_check import ListOfFiguresNotUpdatedCheck
from checks.word.objects.missing_list_of_fugures_check import MissingListOfFiguresCheck
from checks.word.objects.object_caption_bindings_check import ObjectCaptionBindingCheck
from checks.word.objects.object_caption_check import ObjectCaptionCheck
from checks.word.objects.object_caption_description_check import ObjectCaptionDescriptionCheck
from checks.word.objects.object_cross_reference_check import ObjectCrossReferenceCheck
from checks.word.structure.chapter_numbering_continuity_check import ChapterNumberingContinuityCheck
from checks.word.structure.document_structure_check import DocumentStructureCheck
from checks.word.structure.first_chapter_page1_check import FirstChapterStartsOnPageOneCheck
from checks.word.structure.toc_exists_check import TOCExistsCheck
from checks.word.structure.toc_first_section_check import TOCFirstSectionContentCheck
from checks.word.structure.toc_heading_levels_check import TOCHeadingLevelsCheck
from checks.word.structure.toc_illegal_content_check import TOCIllegalContentCheck
from checks.word.structure.toc_up_to_date_check import TOCUpToDateCheck
from document.excel_document import ExcelDocument
from document.word_document import WordDocument
from core.runner import Runner
from core.report import Report

from checks.word.sections.section_count_check import SectionCountCheck
from checks.word.sections.section1_toc_check import Section1TOCCheck
from checks.word.sections.section2_text_check import Section2TextCheck
from checks.word.sections.section3_figure_list_check import Section3FigureListCheck
from checks.word.sections.section3_table_list_check import Section3TableListCheck
from checks.word.sections.section3_bibliography_check import Section3BibliographyCheck

from assignment.word.word_assignment_loader import load_assignment

from checks.word.formatting.normal_style_check import NormalStyleCheck

def main():
    doc = WordDocument("student.docx")
    assignment = load_assignment("assignment/word/assignment.json")
    doc.save_xml()

    checks = [
        # # Části / oddíly 
        # SectionCountCheck(),
        # Section1TOCCheck(),
        # Section2TextCheck(),
        # Section3FigureListCheck(),
        # Section3TableListCheck(),
        # Section3BibliographyCheck(),

        # # Formátování 
        # NormalStyleCheck(),
        # HeadingStyleCheck(1),
        # HeadingStyleCheck(2),
        # HeadingStyleCheck(3),
        # HeadingHierarchicalNumberingCheck(),
        # TocHeadingNumberingCheck(),
        # UnnumberedSpecialHeadingsCheck(),
        # CoverStylesCheck(),
        # FrontpageStylesCheck(),
        # BibliographyStyleCheck(),
        # CaptionStyleCheck(),
        # ContentHeadingStyleCheck(),
        # HeadingsUsedCorrectlyCheck(),
        # OriginalFormattingCheck(),
        # CustomStyleInheritanceCheck(),
        # RequiredCustomStylesUsageCheck(),
        # CustomStyleWithTabsCheck(),
        # MainChapterStartsOnNewPageCheck(),
        # ManualHorizontalSpacingCheck(),
        # ManualVerticalSpacingCheck(),
        # ListLevel2UsedCheck(),
        # InconsistentFormattingCheck(),
        # # Obsah
        # TOCExistsCheck(),
        # TOCUpToDateCheck(),
        # DocumentStructureCheck(),
        # TOCHeadingLevelsCheck(),
        # TOCFirstSectionContentCheck(),
        # TOCIllegalContentCheck(),
        # FirstChapterStartsOnPageOneCheck(),
        # ChapterNumberingContinuityCheck(),
        # # Objekty
        # MissingListOfFiguresCheck(),
        # ListOfFiguresNotUpdatedCheck(),
        # ImageLowQualityCheck(),
        # ObjectCaptionCheck(),
        # ObjectCaptionDescriptionCheck(),
        # ObjectCrossReferenceCheck(),
        # ObjectCaptionBindingCheck(),
        # # Liteatura
        # MissingBibliographyCheck(),
        # BibliographyNotUpdatedCheck(),

        # # Header-Foooter
        # HeaderFooterMissingCheck(),
        # SecondSectionHeaderHasTextCheck(),
        # SecondSectionPageNumberStartsAtOneCheck(),
        # HeaderNotLinkedToPreviousCheck(2),
        # HeaderNotLinkedToPreviousCheck(3),
        # FooterLinkedToPreviousCheck(2),
        # FooterLinkedToPreviousCheck(3),
        # SectionHeaderEmptyCheck(1),
        # SectionHeaderEmptyCheck(3),
        # SectionFooterEmptyCheck(1),
        # SectionFooterHasPageNumberCheck(2),
        # SectionFooterHasPageNumberCheck(3),
    ]

    excel = ExcelDocument("23_fb750.xlsx")
    excel_assignment = load_excel_assignment("assignment/excel/assignment.json")
    excel.save_xml()

    excel_checks = [
        # RequiredSourceWorksheetCheck(),
        # RequiredDataWorksheetCheck(),
        # NonCopyableFormulasCheck(),
        # MissingOrWrongFormulaOrNotCalculatedCheck(),
        # ArrayFormulaCheck(),
        # NamedRangeUsageCheck(),
        # RedundantAbsoluteReferenceCheck(),
        # DescriptiveStatisticsCheck(),

        # formatovani
        # NumberFormattingCheck(),
        # TableBorderCheck(),
        # MergedCellsCheck(),
        # HeaderFormattingCheck(),
        # ConditionalFormattingExistsCheck(),
        # ConditionalFormattingCorrectnessCheck(),
        MissingChartCheck()

    ]

    report = Report()

    word_runner = Runner(checks)
    results = word_runner.run(doc, assignment)

    for check, result in results:
        report.add(check.name, result)

    excel_runner = Runner(excel_checks)
    excel_results = excel_runner.run(excel, excel_assignment)

    for check, result in excel_results:
        report.add(check.name, result)

    report.print()

if __name__ == "__main__":
    main()