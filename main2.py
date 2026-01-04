from pathlib import Path

from assignment.word.word_assignment_loader import load_assignment
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
from checks.word.formatting.normal_style_check import NormalStyleCheck
from checks.word.formatting.original_formatting_check import OriginalFormattingCheck
from checks.word.formatting.toc_heading_numbering_check import TocHeadingNumberingCheck
from checks.word.formatting.unnumbered_special_headings_check import UnnumberedSpecialHeadingsCheck
from checks.word.header_footer.header_footer_missing_check import HeaderFooterMissingCheck
from checks.word.header_footer.second_section_header_text_check import SecondSectionHeaderHasTextCheck
from checks.word.header_footer.second_section_page_num_start_at_one_check import SecondSectionPageNumberStartsAtOneCheck
from checks.word.header_footer.section_emty_footer_check import SectionFooterEmptyCheck
from checks.word.header_footer.section_emty_header_check import SectionHeaderEmptyCheck
from checks.word.header_footer.section_footer_linked_check import FooterLinkedToPreviousCheck
from checks.word.header_footer.section_footer_page_number_check import SectionFooterHasPageNumberCheck
from checks.word.header_footer.section_header_linked_check import HeaderNotLinkedToPreviousCheck
from checks.word.objects.image_low_quality_check import ImageLowQualityCheck
from checks.word.objects.list_of_figures_not_up_to_date_check import ListOfFiguresNotUpdatedCheck
from checks.word.objects.missing_list_of_fugures_check import MissingListOfFiguresCheck
from checks.word.objects.object_caption_bindings_check import ObjectCaptionBindingCheck
from checks.word.objects.object_caption_check import ObjectCaptionCheck
from checks.word.objects.object_caption_description_check import ObjectCaptionDescriptionCheck
from checks.word.objects.object_cross_reference_check import ObjectCrossReferenceCheck
from checks.word.sections.section1_toc_check import Section1TOCCheck
from checks.word.sections.section2_text_check import Section2TextCheck
from checks.word.sections.section3_bibliography_check import Section3BibliographyCheck
from checks.word.sections.section3_figure_list_check import Section3FigureListCheck
from checks.word.sections.section3_table_list_check import Section3TableListCheck
from checks.word.sections.section_count_check import SectionCountCheck
from checks.word.structure.chapter_numbering_continuity_check import ChapterNumberingContinuityCheck
from checks.word.structure.document_structure_check import DocumentStructureCheck
from checks.word.structure.first_chapter_page1_check import FirstChapterStartsOnPageOneCheck
from checks.word.structure.toc_exists_check import TOCExistsCheck
from checks.word.structure.toc_first_section_check import TOCFirstSectionContentCheck
from checks.word.structure.toc_heading_levels_check import TOCHeadingLevelsCheck
from checks.word.structure.toc_illegal_content_check import TOCIllegalContentCheck
from checks.word.structure.toc_up_to_date_check import TOCUpToDateCheck
from core.report import Report
from core.runner import Runner
from document.word_document import WordDocument

# ... tvoje importy ...

def run_checks(word_path: str | Path,
               word_assignment_path: str | Path = "assignment/word/assignment.json",
               run_excel: bool = False,
               excel_path: str | Path = "23_fb750.xlsx",
               excel_assignment_path: str | Path = "assignment/excel/assignment.json"):
    word_path = Path(word_path)
    word_assignment_path = Path(word_assignment_path)

    doc = WordDocument(str(word_path))
    assignment = load_assignment(str(word_assignment_path))
    doc.save_xml()

    checks = [
        SectionCountCheck(),
        Section1TOCCheck(),
        Section2TextCheck(),
        Section3FigureListCheck(),
        Section3TableListCheck(),
        Section3BibliographyCheck(),
        NormalStyleCheck(),
        HeadingStyleCheck(1),
        HeadingStyleCheck(2),
        HeadingStyleCheck(3),
        HeadingHierarchicalNumberingCheck(),
        TocHeadingNumberingCheck(),
        UnnumberedSpecialHeadingsCheck(),
        CoverStylesCheck(),
        FrontpageStylesCheck(),
        BibliographyStyleCheck(),
        CaptionStyleCheck(),
        ContentHeadingStyleCheck(),
        HeadingsUsedCorrectlyCheck(),
        OriginalFormattingCheck(),
        CustomStyleInheritanceCheck(),
        RequiredCustomStylesUsageCheck(),
        CustomStyleWithTabsCheck(),
        MainChapterStartsOnNewPageCheck(),
        ManualHorizontalSpacingCheck(),
        ManualVerticalSpacingCheck(),
        ListLevel2UsedCheck(),
        InconsistentFormattingCheck(),
        TOCExistsCheck(),
        TOCUpToDateCheck(),
        DocumentStructureCheck(),
        TOCHeadingLevelsCheck(),
        TOCFirstSectionContentCheck(),
        TOCIllegalContentCheck(),
        FirstChapterStartsOnPageOneCheck(),
        ChapterNumberingContinuityCheck(),
        MissingListOfFiguresCheck(),
        ListOfFiguresNotUpdatedCheck(),
        ImageLowQualityCheck(),
        ObjectCaptionCheck(),
        ObjectCaptionDescriptionCheck(),
        ObjectCrossReferenceCheck(),
        ObjectCaptionBindingCheck(),
        HeaderFooterMissingCheck(),
        SecondSectionHeaderHasTextCheck(),
        SecondSectionPageNumberStartsAtOneCheck(),
        HeaderNotLinkedToPreviousCheck(2),
        HeaderNotLinkedToPreviousCheck(3),
        FooterLinkedToPreviousCheck(2),
        FooterLinkedToPreviousCheck(3),
        SectionHeaderEmptyCheck(1),
        SectionHeaderEmptyCheck(3),
        SectionFooterEmptyCheck(1),
        SectionFooterHasPageNumberCheck(2),
        SectionFooterHasPageNumberCheck(3),
    ]

    runner = Runner(checks)
    report = Report()

    results = runner.run(doc, assignment)
    for check, result in results:
        report.add(check.name, result)

    return report


from pathlib import Path

def run_folder(docs_dir: str | Path = "docs",
               pattern: str = "*.docx",
               assignment_path: str | Path = "assignment/word/assignment.json"):
    docs_dir = Path(docs_dir)
    files = sorted(docs_dir.glob(pattern))

    if not files:
        raise FileNotFoundError(f"Žádné soubory {pattern} ve složce {docs_dir}")

    for f in files:
        print(f"\n===== Kontroluji: {f.name} =====")
        report = run_checks(f, word_assignment_path=assignment_path, run_excel=False)
        report.print()

def main():
    report = run_checks("student.docx", run_excel=True)
    report.print()


if __name__ == "__main__":
    main()