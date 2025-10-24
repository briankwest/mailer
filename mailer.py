#!/usr/bin/env python3
"""
Unified Letter Generation System for #10 Windowed Envelopes
Generates professional letters with precise fold lines and window alignment
"""

import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Dict, Any, Union
from io import BytesIO

import click
from PIL import Image
from pydantic import BaseModel, Field, field_validator
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch, mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Frame, KeepTogether


# ============================================================================
# DATA MODELS (Pydantic)
# ============================================================================

class Metadata(BaseModel):
    """Letter metadata configuration"""
    type: str = Field(default="business", pattern="^(formal|business|congressional|personal)$")
    date: Optional[str] = None
    date_format: str = Field(default="full", pattern="^(full|abbreviated|custom)$")
    reference_id: str = Field(default="letter_001")

    @field_validator('date')
    @classmethod
    def validate_date(cls, v):
        if v:
            try:
                datetime.strptime(v, "%Y-%m-%d")
            except ValueError:
                raise ValueError("Date must be in YYYY-MM-DD format")
        return v


class Margins(BaseModel):
    """Page margin configuration"""
    top: float = 1.25
    bottom: float = 1.25
    left: float = 1.25
    right: float = 1.25


class AddressPosition(BaseModel):
    """Address positioning for window alignment"""
    x: float
    y: float
    width: float
    height: Optional[float] = None


class DatePosition(BaseModel):
    """Date positioning configuration"""
    x: float = 0.5
    y: float = 1.75
    alignment: str = Field(default="left", pattern="^(left|center|right)$")


class Positioning(BaseModel):
    """Element positioning configuration"""
    unit: str = "inches"
    margins: Margins = Margins()
    return_address: AddressPosition
    recipient_address: AddressPosition
    date_position: DatePosition = DatePosition()
    body_start_y: float = 3.5


class Address(BaseModel):
    """Address structure"""
    name: str
    title: Optional[str] = None
    organization: Optional[str] = None
    street_1: str
    street_2: Optional[str] = None
    city: str
    state: str
    zip: str
    country: Optional[str] = "USA"
    phone: Optional[str] = None
    email: Optional[str] = None
    honorific: Optional[str] = None  # For recipients


class SignatureConfig(BaseModel):
    """Signature configuration"""
    type: str = Field(default="typed", pattern="^(image|typed)$")
    image_path: Optional[str] = None
    width: float = 2.0
    height: float = 0.75
    typed_name: str
    title: Optional[str] = None


class Content(BaseModel):
    """Letter content configuration"""
    salutation: str
    subject: Optional[str] = None
    body: List[str]
    closing: str
    signature: SignatureConfig
    postscript: Optional[str] = None
    enclosures: Optional[List[str]] = None
    cc: Optional[List[str]] = None


class Formatting(BaseModel):
    """Text formatting configuration"""
    font_family: str = "Times-Roman"
    font_size: int = 12
    line_spacing: float = 1.5
    paragraph_spacing: int = 12
    justify_body: bool = False
    indent_paragraphs: bool = True
    indent_size: float = 0.5


class FoldLineStyle(BaseModel):
    """Fold line style configuration"""
    line_length_mm: int = 4
    margin_offset_mm: int = 3
    color: str = "#CCCCCC"
    line_width: float = 0.5
    line_style: str = Field(default="solid", pattern="^(solid|dashed)$")


class FoldLines(BaseModel):
    """Fold lines configuration"""
    enabled: bool = True
    positions: List[float] = [3.67, 7.33]
    style: FoldLineStyle = FoldLineStyle()


class HeaderContent(BaseModel):
    """Header content configuration"""
    enabled: bool = True
    left: str = ""
    center: str = ""
    right: str = ""


class Header(BaseModel):
    """Header configuration"""
    page_1: HeaderContent
    subsequent: HeaderContent
    font_size: int = 10
    color: str = "#333333"
    line_below: bool = True


class Footer(BaseModel):
    """Footer configuration"""
    enabled: bool = True
    left: str = ""
    center: str = "Page {page} of {total}"
    right: str = ""
    font_size: int = 10
    color: str = "#666666"
    line_above: bool = True


class PageNumbers(BaseModel):
    """Page numbering configuration"""
    show: bool = True
    position: str = Field(default="bottom_center", pattern="^(bottom_left|bottom_center|bottom_right)$")
    start_on_page: int = 1


class PageSettings(BaseModel):
    """Page settings configuration"""
    paper_size: str = "letter"
    orientation: str = "portrait"
    page_numbers: PageNumbers = PageNumbers()


class LetterConfig(BaseModel):
    """Complete letter configuration"""
    metadata: Metadata
    positioning: Positioning
    return_address: Address
    recipient_address: Address
    content: Content
    formatting: Formatting
    fold_lines: FoldLines = FoldLines()
    header: Header
    footer: Footer
    page_settings: PageSettings = PageSettings()


# ============================================================================
# PDF GENERATION ENGINE
# ============================================================================

class LetterPDFBuilder:
    """Unified PDF generation engine for letters"""

    def __init__(self, config: LetterConfig):
        self.config = config
        self.buffer = BytesIO()
        self.canvas = None
        self.page_count = 0
        self.total_pages = 0  # Will be calculated
        self.current_y = 0
        self.page_width, self.page_height = letter
        self.fonts_config = self._load_fonts_config()

    def _load_fonts_config(self) -> Dict:
        """Load font configuration"""
        fonts_file = Path(__file__).parent / "config" / "fonts.json"
        if fonts_file.exists():
            with open(fonts_file, 'r') as f:
                return json.load(f)
        return {"default": "Times-Roman", "fallback": "Helvetica"}

    def generate(self) -> bytes:
        """Generate the PDF letter"""
        # Initialize canvas
        self.canvas = canvas.Canvas(self.buffer, pagesize=letter)
        self._set_document_properties()

        # First pass: calculate total pages needed
        self.total_pages = self._calculate_total_pages()

        # Reset for actual generation
        self.buffer = BytesIO()
        self.canvas = canvas.Canvas(self.buffer, pagesize=letter)
        self._set_document_properties()
        self.page_count = 0

        # Generate all pages
        self._generate_pages()

        # Check if actual page count differs from calculated
        # This can happen due to dynamic content in additional elements
        if self.page_count > self.total_pages:
            # Update total pages to match actual count
            self.total_pages = self.page_count

            # Regenerate with correct page count in footers
            self.buffer = BytesIO()
            self.canvas = canvas.Canvas(self.buffer, pagesize=letter)
            self._set_document_properties()
            self.page_count = 0
            self._generate_pages()

        # Save and return PDF
        self.canvas.save()
        pdf = self.buffer.getvalue()
        self.buffer.close()
        return pdf

    def _set_document_properties(self):
        """Set PDF document properties"""
        self.canvas.setTitle(f"Letter - {self.config.metadata.reference_id}")
        self.canvas.setAuthor(self.config.return_address.name)
        self.canvas.setSubject(self.config.content.subject or "Letter")

    def _calculate_total_pages(self) -> int:
        """Calculate total pages needed for the content - accounts for orphan prevention"""
        # Do a dry run to count actual pages needed
        temp_page_count = 1

        # Calculate first page content
        current_y = self.page_height - (self.config.positioning.body_start_y * inch)

        # Account for salutation and subject
        if self.config.content.subject:
            # Subject line plus extra spacing after it
            current_y -= self.config.formatting.paragraph_spacing * 1.5
        current_y -= self.config.formatting.font_size * self.config.formatting.line_spacing * 1.5

        # Calculate space needed for body paragraphs
        max_width = (self.page_width - self.config.positioning.margins.left * inch -
                    self.config.positioning.margins.right * inch)
        bottom_margin = 0.75 * inch  # Use same bottom margin as _flow_body_text
        page_third = (self.page_height - self.config.positioning.margins.top * inch) / 3

        paragraphs = self.config.content.body
        i = 0
        while i < len(paragraphs):
            paragraph = paragraphs[i]

            # Check if this is likely a heading (same logic as _flow_body_text)
            is_heading = (paragraph.isupper() and
                         len(paragraph.split()) <= 10 and
                         not paragraph.rstrip().endswith(('.', '!', '?', ',')))

            # Wrap text to get actual lines
            lines = self._wrap_text(paragraph, max_width - (self.config.formatting.indent_size * inch if self.config.formatting.indent_paragraphs else 0))

            # Calculate space needed for entire paragraph
            space_needed = len(lines) * self.config.formatting.font_size * self.config.formatting.line_spacing
            if i < len(paragraphs) - 1:
                space_needed += self.config.formatting.paragraph_spacing

            # If this is a heading, apply orphan prevention logic
            if is_heading and i < len(paragraphs) - 1:
                # Get the next paragraph and calculate its space requirements
                next_para = paragraphs[i + 1]
                next_lines = self._wrap_text(next_para, max_width - (self.config.formatting.indent_size * inch if self.config.formatting.indent_paragraphs else 0))

                # We want at least 4 lines of the next paragraph to appear with the heading
                min_next_lines = min(4, max(2, len(next_lines) // 2))
                space_for_next = min_next_lines * self.config.formatting.font_size * self.config.formatting.line_spacing

                # Total space needed is heading + paragraph spacing + at least 4 lines of next paragraph
                min_space_needed = space_needed + space_for_next

                # Check both conditions: space needed AND bottom third rule
                if current_y < bottom_margin + min_space_needed or current_y < page_third:
                    # Move heading to next page
                    temp_page_count += 1
                    current_y = self.page_height - (self.config.positioning.margins.top * inch)

            # Check if paragraph fits on current page
            if current_y < bottom_margin + space_needed:
                temp_page_count += 1
                current_y = self.page_height - (self.config.positioning.margins.top * inch)

            # Account for the space used by this paragraph
            current_y -= space_needed
            i += 1

        # Account for closing and signature (with proper space calculation)
        signature_space = 3 * inch  # Space for closing, signature area, and name/title
        if current_y < bottom_margin + signature_space:
            temp_page_count += 1
            current_y = self.page_height - (self.config.positioning.margins.top * inch)

        # Account for additional elements (P.S., enclosures, CC)
        additional_space = 0
        if self.config.content.postscript:
            ps_lines = self._wrap_text(f"P.S. {self.config.content.postscript}", max_width)
            additional_space += len(ps_lines) * self.config.formatting.font_size * 1.2 + self.config.formatting.paragraph_spacing

        if self.config.content.enclosures:
            additional_space += (len(self.config.content.enclosures) + 1) * self.config.formatting.font_size * 1.2

        if self.config.content.cc:
            additional_space += (len(self.config.content.cc) + 1) * self.config.formatting.font_size * 1.2

        if additional_space > 0 and current_y - additional_space < bottom_margin:
            temp_page_count += 1

        return temp_page_count

    def _generate_pages(self):
        """Generate all pages of the letter"""
        self.page_count = 1

        # Page 1
        self._start_new_page()
        self._draw_return_address()
        self._draw_date()
        self._draw_recipient_address()
        self._draw_salutation()

        # Flow body text
        remaining_body = self._flow_body_text()

        # Continue on additional pages if needed
        while remaining_body:
            self.page_count += 1  # Increment page count BEFORE starting new page
            self._start_new_page()
            remaining_body = self._flow_body_text(remaining_body)

        # Add closing and signature
        self._draw_closing_signature()

        # Add postscript, enclosures, cc if present
        self._draw_additional_elements()

    def _start_new_page(self):
        """Start a new page"""
        if self.page_count > 1:
            self.canvas.showPage()

        # Draw page elements
        self._draw_fold_lines()
        self._draw_header()
        self._draw_footer()

        # Set starting Y position
        if self.page_count == 1:
            self.current_y = self.page_height - (self.config.positioning.body_start_y * inch)
        else:
            self.current_y = self.page_height - (self.config.positioning.margins.top * inch)

    def _draw_fold_lines(self):
        """Draw fold lines in margins"""
        if not self.config.fold_lines.enabled:
            return

        # Convert color from hex to RGB
        color_hex = self.config.fold_lines.style.color.lstrip('#')
        r, g, b = tuple(int(color_hex[i:i+2], 16)/255 for i in (0, 2, 4))
        self.canvas.setStrokeColor(colors.Color(r, g, b))
        self.canvas.setLineWidth(self.config.fold_lines.style.line_width)

        # Draw lines at fold positions
        for fold_y in self.config.fold_lines.positions:
            y_pos = self.page_height - (fold_y * inch)

            # Left margin line
            left_x = self.config.fold_lines.style.margin_offset_mm * mm
            line_length = self.config.fold_lines.style.line_length_mm * mm
            self.canvas.line(left_x, y_pos, left_x + line_length, y_pos)

            # Right margin line
            right_x = self.page_width - left_x
            self.canvas.line(right_x - line_length, y_pos, right_x, y_pos)

    def _draw_header(self):
        """Draw page header"""
        if self.page_count == 1:
            header_config = self.config.header.page_1
            # Skip header on page 1 by default
            if not header_config.enabled:
                return
        else:
            header_config = self.config.header.subsequent

        if not header_config.enabled:
            return

        # Set font and color
        self.canvas.setFont(self.config.formatting.font_family, self.config.header.font_size)
        color_hex = self.config.header.color.lstrip('#')
        r, g, b = tuple(int(color_hex[i:i+2], 16)/255 for i in (0, 2, 4))
        self.canvas.setFillColor(colors.Color(r, g, b))

        # Calculate positions
        y_pos = self.page_height - (0.5 * inch)
        left_x = self.config.positioning.margins.left * inch
        right_x = self.page_width - (self.config.positioning.margins.right * inch)
        center_x = self.page_width / 2

        # Format and draw header content
        formatted_date = self._format_date()

        # Left content
        left_text = header_config.left.format(
            return_name=self.config.return_address.name,
            recipient_name=self.config.recipient_address.name,
            formatted_date=formatted_date,
            page=self.page_count
        )
        if left_text:
            self.canvas.drawString(left_x, y_pos, left_text)

        # Center content
        center_text = header_config.center.format(
            page=self.page_count,
            formatted_date=formatted_date
        )
        if center_text:
            text_width = stringWidth(center_text, self.config.formatting.font_family,
                                    self.config.header.font_size)
            self.canvas.drawString(center_x - text_width/2, y_pos, center_text)

        # Right content
        right_text = header_config.right.format(
            formatted_date=formatted_date,
            page=self.page_count
        )
        if right_text:
            text_width = stringWidth(right_text, self.config.formatting.font_family,
                                   self.config.header.font_size)
            self.canvas.drawString(right_x - text_width, y_pos, right_text)

        # Draw line below header
        if self.config.header.line_below:
            self.canvas.setStrokeColor(colors.Color(0.8, 0.8, 0.8))
            self.canvas.setLineWidth(0.5)
            y_line = y_pos - 5
            self.canvas.line(left_x, y_line, right_x, y_line)

        # Reset to black for body text
        self.canvas.setFillColor(colors.black)
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)

    def _draw_footer(self):
        """Draw page footer"""
        if not self.config.footer.enabled:
            return

        # Set font and color
        self.canvas.setFont(self.config.formatting.font_family, self.config.footer.font_size)
        color_hex = self.config.footer.color.lstrip('#')
        r, g, b = tuple(int(color_hex[i:i+2], 16)/255 for i in (0, 2, 4))
        self.canvas.setFillColor(colors.Color(r, g, b))

        # Calculate positions
        y_pos = 0.5 * inch
        left_x = self.config.positioning.margins.left * inch
        right_x = self.page_width - (self.config.positioning.margins.right * inch)
        center_x = self.page_width / 2

        # Draw line above footer
        if self.config.footer.line_above:
            self.canvas.setStrokeColor(colors.Color(0.8, 0.8, 0.8))
            self.canvas.setLineWidth(0.5)
            y_line = y_pos + 15
            self.canvas.line(left_x, y_line, right_x, y_line)

        # Center content (page numbers)
        center_text = self.config.footer.center.format(
            page=self.page_count,
            total=self.total_pages
        )
        if center_text:
            text_width = stringWidth(center_text, self.config.formatting.font_family,
                                   self.config.footer.font_size)
            self.canvas.drawString(center_x - text_width/2, y_pos, center_text)

        # Reset to black for body text
        self.canvas.setFillColor(colors.black)
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)

    def _format_date(self) -> str:
        """Format date according to configuration"""
        if self.config.metadata.date:
            date_obj = datetime.strptime(self.config.metadata.date, "%Y-%m-%d")
        else:
            date_obj = datetime.now()

        if self.config.metadata.date_format == "full":
            # Format as "October 24, 2025"
            return date_obj.strftime("%B %d, %Y").replace(" 0", " ")
        elif self.config.metadata.date_format == "abbreviated":
            # Format as "Oct 24, 2025"
            return date_obj.strftime("%b %d, %Y").replace(" 0", " ")
        else:
            # Custom format
            return date_obj.strftime("%Y-%m-%d")

    def _draw_return_address(self):
        """Draw return address in window position"""
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)
        self.canvas.setFillColor(colors.black)

        x = self.config.positioning.return_address.x * inch
        y = self.page_height - (self.config.positioning.return_address.y * inch)

        # Build address lines
        lines = []
        if self.config.return_address.name:
            lines.append(self.config.return_address.name)
        if self.config.return_address.title:
            lines.append(self.config.return_address.title)
        if self.config.return_address.organization:
            lines.append(self.config.return_address.organization)
        lines.append(self.config.return_address.street_1)
        if self.config.return_address.street_2:
            lines.append(self.config.return_address.street_2)
        lines.append(f"{self.config.return_address.city}, {self.config.return_address.state} {self.config.return_address.zip}")

        # Draw each line
        for line in lines:
            self.canvas.drawString(x, y, line)
            y -= self.config.formatting.font_size * 1.2

    def _draw_date(self):
        """Draw date - positioned properly between addresses"""
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)
        self.canvas.setFillColor(colors.black)  # Ensure date is black

        y = self.page_height - (self.config.positioning.date_position.y * inch)
        formatted_date = self._format_date()

        # Handle different alignments
        if self.config.positioning.date_position.alignment == "right":
            # Right align - align to right edge of recipient address block
            # Recipient address ends at x=0.75" + width=4.0" = 4.75"
            x = (self.config.positioning.recipient_address.x + self.config.positioning.recipient_address.width) * inch
            text_width = stringWidth(formatted_date, self.config.formatting.font_family,
                                    self.config.formatting.font_size)
            self.canvas.drawString(x - text_width, y, formatted_date)
        elif self.config.positioning.date_position.alignment == "center":
            # Center on page
            x = self.page_width / 2
            text_width = stringWidth(formatted_date, self.config.formatting.font_family,
                                    self.config.formatting.font_size)
            self.canvas.drawString(x - text_width/2, y, formatted_date)
        else:
            # Left align (default)
            x = self.config.positioning.date_position.x * inch
            self.canvas.drawString(x, y, formatted_date)

    def _draw_recipient_address(self):
        """Draw recipient address in window position"""
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)

        x = self.config.positioning.recipient_address.x * inch
        y = self.page_height - (self.config.positioning.recipient_address.y * inch)

        # Build address lines
        lines = []
        if self.config.recipient_address.honorific:
            lines.append(f"{self.config.recipient_address.honorific} {self.config.recipient_address.name}")
        else:
            lines.append(self.config.recipient_address.name)

        if self.config.recipient_address.title:
            lines.append(self.config.recipient_address.title)
        if self.config.recipient_address.organization:
            lines.append(self.config.recipient_address.organization)
        lines.append(self.config.recipient_address.street_1)
        if self.config.recipient_address.street_2:
            lines.append(self.config.recipient_address.street_2)
        lines.append(f"{self.config.recipient_address.city}, {self.config.recipient_address.state} {self.config.recipient_address.zip}")

        # Draw each line
        for line in lines:
            self.canvas.drawString(x, y, line)
            y -= self.config.formatting.font_size * 1.2

    def _draw_salutation(self):
        """Draw salutation"""
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)

        x = self.config.positioning.margins.left * inch

        # Add spacing before salutation
        self.current_y -= self.config.formatting.paragraph_spacing

        # Draw subject if present
        if self.config.content.subject:
            # Use the correct bold font name
            if "Times" in self.config.formatting.font_family:
                bold_font = "Times-Bold"
            elif "Helvetica" in self.config.formatting.font_family:
                bold_font = "Helvetica-Bold"
            elif "Courier" in self.config.formatting.font_family:
                bold_font = "Courier-Bold"
            else:
                bold_font = "Helvetica-Bold"  # Fallback

            self.canvas.setFont(bold_font, self.config.formatting.font_size)
            self.canvas.drawString(x, self.current_y, self.config.content.subject)
            # Add extra spacing after subject (paragraph spacing instead of just line spacing)
            self.current_y -= self.config.formatting.paragraph_spacing * 1.5
            self.canvas.setFont(self.config.formatting.font_family,
                              self.config.formatting.font_size)

        # Draw salutation
        self.canvas.drawString(x, self.current_y, f"{self.config.content.salutation}:")
        self.current_y -= self.config.formatting.font_size * self.config.formatting.line_spacing * 1.5

    def _flow_body_text(self, remaining_text: Optional[List[str]] = None) -> List[str]:
        """Flow body text with multi-page support - keeps paragraphs together and prevents orphaned headings"""
        paragraphs = remaining_text if remaining_text else self.config.content.body

        x = self.config.positioning.margins.left * inch
        max_width = (self.page_width -
                    (self.config.positioning.margins.left + self.config.positioning.margins.right) * inch)
        # Use space down to footer line (about 0.75 inch from bottom for footer)
        bottom_margin = 0.75 * inch

        remaining = []

        for i, paragraph in enumerate(paragraphs):
            # Check if this is likely a heading (all caps, short, no ending punctuation)
            is_heading = (paragraph.isupper() and
                         len(paragraph.split()) <= 10 and
                         not paragraph.rstrip().endswith(('.', '!', '?', ',')))

            # Wrap text first to know how many lines we need
            lines = self._wrap_text(paragraph, max_width - (self.config.formatting.indent_size * inch if self.config.formatting.indent_paragraphs else 0))

            # Calculate space needed for this entire paragraph
            lines_needed = len(lines)
            space_needed = lines_needed * self.config.formatting.font_size * self.config.formatting.line_spacing

            # Add paragraph spacing if not the last paragraph
            if i < len(paragraphs) - 1:
                space_needed += self.config.formatting.paragraph_spacing

            # If this is a heading, check if we have room for it plus at least 4 lines of the next paragraph
            if is_heading and i < len(paragraphs) - 1:
                # Get the next paragraph and calculate its space requirements
                next_para = paragraphs[i + 1]
                next_lines = self._wrap_text(next_para, max_width - (self.config.formatting.indent_size * inch if self.config.formatting.indent_paragraphs else 0))

                # We want at least 4 lines of the next paragraph to appear with the heading
                # or half the paragraph, whichever is smaller
                min_next_lines = min(4, max(2, len(next_lines) // 2))
                space_for_next = min_next_lines * self.config.formatting.font_size * self.config.formatting.line_spacing

                # Total space needed is heading + paragraph spacing + at least 4 lines of next paragraph
                min_space_needed = space_needed + space_for_next

                # Be more aggressive - if we're in the bottom third of the page, move to next
                page_third = (self.page_height - self.config.positioning.margins.top * inch) / 3

                if self.current_y < bottom_margin + min_space_needed or self.current_y < page_third:
                    # Move heading to next page
                    remaining = paragraphs[i:]
                    break
            # If the entire paragraph doesn't fit, move it to the next page
            elif self.current_y < bottom_margin + space_needed:
                remaining = paragraphs[i:]
                break

            # Add paragraph indent if configured
            para_x = x
            if self.config.formatting.indent_paragraphs and not is_heading:
                # First line of each paragraph gets indented (but not headings)
                first_line_x = para_x + self.config.formatting.indent_size * inch
            else:
                first_line_x = para_x

            # Draw all lines of the paragraph (we know they fit)
            for j, line in enumerate(lines):
                # Use indented position for first line only (not for headings)
                line_x = first_line_x if j == 0 and not is_heading else para_x
                self.canvas.drawString(line_x, self.current_y, line)
                self.current_y -= self.config.formatting.font_size * self.config.formatting.line_spacing

            # Add paragraph spacing - only if not the last paragraph
            if i < len(paragraphs) - 1:
                self.current_y -= self.config.formatting.paragraph_spacing

        return remaining

    def _wrap_text(self, text: str, max_width: float) -> List[str]:
        """Wrap text to fit within max width"""
        words = text.split()
        lines = []
        current_line = []

        for word in words:
            test_line = ' '.join(current_line + [word])
            text_width = stringWidth(test_line, self.config.formatting.font_family,
                                   self.config.formatting.font_size)

            if text_width > max_width and current_line:
                lines.append(' '.join(current_line))
                current_line = [word]
            else:
                current_line.append(word)

        if current_line:
            lines.append(' '.join(current_line))

        return lines

    def _draw_closing_signature(self):
        """Draw closing and signature"""
        # Add spacing before closing (double paragraph spacing)
        self.current_y -= self.config.formatting.paragraph_spacing * 2

        # Check if we need a new page
        required_space = 3 * inch  # Space for closing, signature, and name
        if self.current_y < self.config.positioning.margins.bottom * inch + required_space:
            self.page_count += 1
            self._start_new_page()

        x = self.config.positioning.margins.left * inch

        # Draw closing
        self.canvas.setFont(self.config.formatting.font_family, self.config.formatting.font_size)
        self.canvas.setFillColor(colors.black)  # Ensure text is black
        self.canvas.drawString(x, self.current_y, f"{self.config.content.closing},")
        self.current_y -= 0.25 * inch  # Tighter space for signature

        # Draw signature (image or typed)
        # Always reserve the same space for signatures (for manual signing with pen)
        signature_space = self.config.content.signature.height * inch * 0.8  # Default 0.75" * 0.8 = 0.6"

        if self.config.content.signature.type == "image" and self.config.content.signature.image_path:
            # Load and draw signature image
            sig_path = Path(self.config.content.signature.image_path)
            if sig_path.exists():
                try:
                    img = Image.open(sig_path)
                    img_reader = ImageReader(img)
                    # Draw signature image closer to the closing
                    self.canvas.drawImage(img_reader, x, self.current_y - signature_space,
                                        width=self.config.content.signature.width * inch,
                                        height=self.config.content.signature.height * inch,
                                        preserveAspectRatio=True, mask='auto')
                except Exception as e:
                    print(f"Warning: Could not load signature image: {e}")

        # Move down by the signature space (same for both typed and image)
        self.current_y -= signature_space

        # Add small padding between signature area and typed name
        self.current_y -= 0.15 * inch
        self.canvas.setFillColor(colors.black)  # Ensure text is black
        self.canvas.drawString(x, self.current_y, self.config.content.signature.typed_name)

        # Draw title if present
        if self.config.content.signature.title:
            self.current_y -= self.config.formatting.font_size * 1.2
            self.canvas.drawString(x, self.current_y, self.config.content.signature.title)

    def _draw_additional_elements(self):
        """Draw postscript, enclosures, and CC list"""
        x = self.config.positioning.margins.left * inch
        max_width = (self.page_width -
                    (self.config.positioning.margins.left + self.config.positioning.margins.right) * inch)
        bottom_margin = 0.75 * inch

        # Postscript
        if self.config.content.postscript:
            # Calculate space needed
            ps_lines = self._wrap_text(f"P.S. {self.config.content.postscript}", max_width)
            space_needed = (self.config.formatting.paragraph_spacing * 2 +
                          len(ps_lines) * self.config.formatting.font_size * self.config.formatting.line_spacing)

            # Check if we need a new page
            if self.current_y < bottom_margin + space_needed:
                self.page_count += 1
                self._start_new_page()

            self.current_y -= self.config.formatting.paragraph_spacing * 2
            self.canvas.drawString(x, self.current_y, "P.S. " + ps_lines[0] if ps_lines else "")
            self.current_y -= self.config.formatting.font_size * self.config.formatting.line_spacing

            # Draw remaining lines
            for line in ps_lines[1:]:
                self.canvas.drawString(x, self.current_y, line)
                self.current_y -= self.config.formatting.font_size * self.config.formatting.line_spacing

        # Enclosures
        if self.config.content.enclosures:
            # Calculate space needed
            space_needed = (self.config.formatting.paragraph_spacing * 2 +
                          self.config.formatting.font_size * 1.2 * (len(self.config.content.enclosures) + 1))

            # Check if we need a new page
            if self.current_y < bottom_margin + space_needed:
                self.page_count += 1
                self._start_new_page()

            self.current_y -= self.config.formatting.paragraph_spacing * 2
            self.canvas.drawString(x, self.current_y, "Enclosures:")
            for enc in self.config.content.enclosures:
                self.current_y -= self.config.formatting.font_size * 1.2
                self.canvas.drawString(x + 0.25 * inch, self.current_y, f"- {enc}")

        # CC list
        if self.config.content.cc:
            # Calculate space needed
            space_needed = (self.config.formatting.paragraph_spacing * 2 +
                          self.config.formatting.font_size * 1.2 * (len(self.config.content.cc) + 1))

            # Check if we need a new page
            if self.current_y < bottom_margin + space_needed:
                self.page_count += 1
                self._start_new_page()

            self.current_y -= self.config.formatting.paragraph_spacing * 2
            self.canvas.drawString(x, self.current_y, "cc:")
            for cc in self.config.content.cc:
                self.current_y -= self.config.formatting.font_size * 1.2
                self.canvas.drawString(x + 0.25 * inch, self.current_y, cc)


# ============================================================================
# MAC PRINTING INTEGRATION
# ============================================================================

class MacPrinter:
    """Mac-specific printing functionality"""

    def __init__(self):
        self.preview_app = "/System/Applications/Preview.app"

    def open_in_preview(self, pdf_path: str):
        """Open PDF in Preview.app"""
        try:
            subprocess.run(["open", "-a", "Preview", pdf_path], check=True)
            return True
        except subprocess.CalledProcessError as e:
            print(f"Error opening Preview: {e}")
            return False

    def print_directly(self, pdf_path: str, printer: Optional[str] = None):
        """Print using lpr command"""
        try:
            cmd = ["lpr"]
            if printer:
                cmd.extend(["-P", printer])
            cmd.append(pdf_path)
            subprocess.run(cmd, check=True)
            return True
        except subprocess.CalledProcessError as e:
            print(f"Error printing: {e}")
            return False

    def get_printers(self) -> List[str]:
        """List available printers"""
        try:
            result = subprocess.run(
                ["lpstat", "-p"],
                capture_output=True,
                text=True,
                check=True
            )
            # Parse printer names from output
            printers = []
            for line in result.stdout.splitlines():
                if line.startswith("printer"):
                    parts = line.split()
                    if len(parts) >= 2:
                        printers.append(parts[1])
            return printers
        except subprocess.CalledProcessError:
            return []

    def print_with_dialog(self, pdf_path: str):
        """Open print dialog via AppleScript"""
        script = f'''
        tell application "Preview"
            open "{pdf_path}"
            activate
            delay 1
            print document 1 with print dialog
        end tell
        '''
        try:
            subprocess.run(["osascript", "-e", script], check=True)
            return True
        except subprocess.CalledProcessError as e:
            print(f"Error opening print dialog: {e}")
            return False


# ============================================================================
# CLI INTERFACE
# ============================================================================

@click.command()
@click.argument('input_json', type=click.Path(exists=True), required=False)
@click.option('--output', '-o', help='Output PDF path')
@click.option('--preview', is_flag=True, help='Open in Preview.app')
@click.option('--print', 'do_print', is_flag=True, help='Send directly to printer')
@click.option('--print-dialog', is_flag=True, help='Open print dialog')
@click.option('--printer', help='Specific printer name')
@click.option('--list-printers', is_flag=True, help='List available printers')
@click.option('--validate', is_flag=True, help='Validate JSON only')
@click.option('--font', help='Override font selection')
def generate_letter(input_json, output, preview, do_print, print_dialog,
                   printer, list_printers, validate, font):
    """
    Generate professional letters for #10 windowed envelopes.

    This tool creates PDF letters with precise formatting for standard
    #10 double-window envelopes, including fold lines and proper address
    positioning.
    """

    # List printers if requested
    if list_printers:
        mac_printer = MacPrinter()
        printers = mac_printer.get_printers()
        if printers:
            click.echo("Available printers:")
            for p in printers:
                click.echo(f"  - {p}")
        else:
            click.echo("No printers found")
        return

    # Check if input JSON is provided for other operations
    if not input_json:
        click.echo("Error: INPUT_JSON is required for letter generation", err=True)
        click.echo("Use --list-printers to list available printers without generating a letter", err=True)
        sys.exit(1)

    # Load and validate JSON configuration
    try:
        with open(input_json, 'r') as f:
            config_dict = json.load(f)

        # Create Pydantic model (validates structure)
        config = LetterConfig(**config_dict)

        # Override font if specified
        if font:
            config.formatting.font_family = font

        if validate:
            click.echo("✓ JSON configuration is valid")
            return

    except FileNotFoundError:
        click.echo(f"✗ Error: File not found: {input_json}", err=True)
        sys.exit(1)
    except json.JSONDecodeError as e:
        click.echo(f"✗ Error: Invalid JSON: {e}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"✗ Error: Validation failed: {e}", err=True)
        sys.exit(1)

    # Generate PDF
    try:
        click.echo(f"Generating letter for {config.recipient_address.name}...")

        pdf_builder = LetterPDFBuilder(config)
        pdf_data = pdf_builder.generate()

        # Determine output path
        if not output:
            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            output = output_dir / f"{config.metadata.reference_id}.pdf"
        else:
            output = Path(output)
            output.parent.mkdir(parents=True, exist_ok=True)

        # Save PDF
        with open(output, 'wb') as f:
            f.write(pdf_data)

        click.echo(f"✓ Letter generated: {output}")

    except Exception as e:
        click.echo(f"✗ Error generating PDF: {e}", err=True)
        sys.exit(1)

    # Handle output actions
    mac_printer = MacPrinter()

    if preview or (not do_print and not print_dialog):
        # Default to preview if no other action specified
        click.echo("Opening in Preview...")
        mac_printer.open_in_preview(str(output))

    if print_dialog:
        click.echo("Opening print dialog...")
        mac_printer.print_with_dialog(str(output))

    elif do_print:
        if printer:
            click.echo(f"Sending to printer: {printer}")
        else:
            click.echo("Sending to default printer...")
        mac_printer.print_directly(str(output), printer)


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    generate_letter()