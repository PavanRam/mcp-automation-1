from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.prompts import base
from mcp.types import TextContent 
import win32com.client
import pywinauto
from pywinauto.application import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto import Desktop
from typing import Optional
import os
import sys
import pythoncom
import platform
import time
import logging
from datetime import datetime
import smtplib
from email.message import EmailMessage
import mimetypes
from dotenv import load_dotenv


# FastMCP initializes the server.
# base.Message structures the prompt for the LLM.
# TextContent defines the specific plain text string within that message

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Global variable to store PowerPoint application instance
ppt_app = None
ppt_window = None


# Instantiate the FastMCP server
mcp = FastMCP("PowerPoint Automation")


# Load environment variables from .env file
load_dotenv()

# DEFINE RESOURCES

# Add a dynamic greeting resource
@mcp.resource("greeting://{name}")
def get_greeting(name: str) -> str:
    """Get a personalized greeting"""
    print("CALLED: get_greeting(name: str) -> str:")
    return f"Hello, {name}!"


# DEFINE AVAILABLE PROMPTS
@mcp.prompt()
def review_code(code: str) -> str:
    return f"Please review this code:\n\n{code}"
    print("CALLED: review_code(code: str) -> str:")


@mcp.prompt()
def debug_error(error: str) -> list[base.Message]:
    return [
        base.UserMessage("I'm seeing this error:"),
        base.UserMessage(error),
        base.AssistantMessage("I'll help debug that. What have you tried so far?"),
    ]
    

    
@mcp.tool()
def open_powerpoint() -> dict:
    """
    Opens Microsoft PowerPoint application and creates a new blank presentation.
    
    Returns:
        dict: Response with content containing success message or error description
    """
    global ppt_app, ppt_presentation
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Close any existing presentation first
        if ppt_app is not None and ppt_presentation is not None:
            try:
                ppt_presentation.Close()
                logger.info("Closed existing presentation")
            except:
                pass
        
        # Try to connect to existing PowerPoint instance or create new one
        try:
            ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            logger.info("Connected to existing PowerPoint instance")
        except:
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            logger.info("Created new PowerPoint instance")
        
        # Make PowerPoint visible
        ppt_app.Visible = True
        
        # Create a new blank presentation (this should only create one)
        ppt_presentation = ppt_app.Presentations.Add(WithWindow=True)
        
        # Add a blank slide (layout 12 is blank)
        slide = ppt_presentation.Slides.Add(1, 12)  # 12 = ppLayoutBlank
        
        logger.info(f"New blank presentation created. Total presentations: {ppt_app.Presentations.Count}")
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"PowerPoint opened successfully with a new blank presentation (Total open: {ppt_app.Presentations.Count})"
                )
            ]
        }
    
    except Exception as e:
        logger.error(f"Error opening PowerPoint: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error opening PowerPoint: {str(e)}"
                )
            ]
        }


@mcp.tool()
def draw_rectangle(x: int = 100, y: int = 100, width: int = 200, height: int = 100) -> dict:
    """
    Draws a rectangle on the current PowerPoint slide.
    
    Args:
        x: X coordinate of the top-left corner in points (default: 100)
        y: Y coordinate of the top-left corner in points (default: 100)
        width: Width of the rectangle in points (default: 200)
        height: Height of the rectangle in points (default: 100)
    
    Returns:
        dict: Response with content containing success message or error description
    """
    global ppt_app, ppt_presentation
    
    if ppt_app is None or ppt_presentation is None:
        return {
            "content": [
                TextContent(
                    type="text",
                    text="Error: PowerPoint is not open. Please run open_powerpoint() first."
                )
            ]
        }
    
    try:
        # Get the current slide
        if ppt_presentation.Slides.Count == 0:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="Error: No slides available. Please add a slide first."
                    )
                ]
            }
        
        slide = ppt_presentation.Slides(ppt_presentation.Slides.Count)
        
        # Add a rectangle shape
        # msoShapeRectangle = 1
        rectangle = slide.Shapes.AddShape(1, float(x), float(y), float(width), float(height))
        
        logger.info(f"Rectangle drawn at ({x}, {y}) with size {width}x{height}")
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Rectangle drawn at ({x}, {y}) with size {width}x{height}"
                )
            ]
        }
    
    except Exception as e:
        logger.error(f"Error drawing rectangle: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error drawing rectangle: {str(e)}"
                )
            ]
        }


@mcp.tool()
def draw_rectangle_with_text(
    text: str,
    x: int = 100,
    y: int = 100,
    width: int = 200,
    height: int = 100
) -> dict:
    """
    Draws a rectangle on the current PowerPoint slide and adds text in the middle.
    
    Args:
        text: Text to display in the center of the rectangle
        x: X coordinate of the top-left corner in points (default: 100)
        y: Y coordinate of the top-left corner in points (default: 100)
        width: Width of the rectangle in points (default: 200)
        height: Height of the rectangle in points (default: 100)
    
    Returns:
        dict: Response with content containing success message or error description
    """
    global ppt_app, ppt_presentation
    
    if ppt_app is None or ppt_presentation is None:
        return {
            "content": [
                TextContent(
                    type="text",
                    text="Error: PowerPoint is not open. Please run open_powerpoint() first."
                )
            ]
        }
    
    try:
        # Get the current slide
        if ppt_presentation.Slides.Count == 0:
            return {
                "content": [
                    TextContent(
                        type="text",
                        text="Error: No slides available. Please add a slide first."
                    )
                ]
            }
        
        slide = ppt_presentation.Slides(ppt_presentation.Slides.Count)
        
        # Add a rectangle shape
        # msoShapeRectangle = 1
        rectangle = slide.Shapes.AddShape(1, float(x), float(y), float(width), float(height))
        
        # Add text to the rectangle
        text_frame = rectangle.TextFrame
        text_frame.TextRange.Text = text
        
        # Center align the text
        text_frame.TextRange.ParagraphFormat.Alignment = 2  # ppAlignCenter = 2
        
        # Vertically center the text
        text_frame.VerticalAnchor = 3  # msoAnchorMiddle = 3
        
        logger.info(f"Rectangle with text '{text}' created at ({x}, {y})")
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Rectangle with text '{text}' created at ({x}, {y}) with size {width}x{height}"
                )
            ]
        }
    
    except Exception as e:
        logger.error(f"Error drawing rectangle with text: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error drawing rectangle with text: {str(e)}"
                )
            ]
        }
        
@mcp.tool()
def save_presentation(filename: str, add_timestamp: bool = True) -> dict:
    """
    Saves the current PowerPoint presentation to a file.
    
    Args:
        filename: Full path where to save the presentation (e.g., "C:\\Users\\YourName\\Documents\\presentation.pptx")
        add_timestamp: If True, adds timestamp to filename to avoid conflicts (default: True)
    
    Returns:
        dict: Response with content containing success message or error description
    """
    global ppt_app, ppt_presentation
    
    if ppt_app is None or ppt_presentation is None:
        return {
            "content": [
                TextContent(
                    type="text",
                    text="Error: PowerPoint is not open. Please run open_powerpoint() first."
                )
            ]
        }
    
    try:
        # Parse the filename path
        directory = os.path.dirname(filename)
        basename = os.path.basename(filename)
        name, ext = os.path.splitext(basename)
        
        # Ensure we have an extension
        if not ext:
            ext = '.pptx'
        elif ext.lower() not in ['.pptx', '.ppt']:
            ext = '.pptx'
        
        # Add timestamp if requested
        if add_timestamp:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            name = f"{name}_{timestamp}"
        
        # Reconstruct the full path
        if directory:
            final_filename = os.path.join(directory, name + ext)
        else:
            final_filename = name + ext
        
        # Save the presentation
        ppt_presentation.SaveAs(final_filename)
        
        logger.info(f"Presentation saved to: {final_filename}")
        print(f"Presentation saved to: {final_filename}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Presentation saved successfully to: {final_filename}"
                )
            ]
        }
    
    except Exception as e:
        logger.error(f"Error saving presentation: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error saving presentation: {str(e)}"
                )
            ]
        }


@mcp.tool()
def close_powerpoint(save: bool = False, filename: str = None) -> dict:
    """
    Closes the PowerPoint application.
    
    Args:
        save: Whether to save the presentation before closing (default: False)
        filename: Optional filename to save as (e.g., "C:\\presentation.pptx")
    
    Returns:
        dict: Response with content containing success message or error description
    """
    global ppt_app, ppt_presentation
    
    if ppt_app is None:
        return {
            "content": [
                TextContent(
                    type="text",
                    text="PowerPoint is not open"
                )
            ]
        }
    
    try:
        if save and ppt_presentation:
            if filename:
                ppt_presentation.SaveAs(filename)
            else:
                ppt_presentation.Save()
        
        if ppt_presentation:
            ppt_presentation.Close()
        
        ppt_app.Quit()
        ppt_app = None
        ppt_presentation = None
        
        return {
            "content": [
                TextContent(
                    type="text",
                    text="PowerPoint closed successfully"
                )
            ]
        }
    
    except Exception as e:
        logger.error(f"Error closing PowerPoint: {str(e)}")
        return {
            "content": [
                TextContent(
                    type="text",
                    text=f"Error closing PowerPoint: {str(e)}"
                )
            ]
        }
        

           
           
if __name__ == "__main__":
    # Check if running with mcp dev command
    print("STARTING")
    if len(sys.argv) > 1 and sys.argv[1] == "dev":
        mcp.run()  # Run without transport for dev server
    else:
        mcp.run(transport="stdio")  # Run with stdio for direct execution