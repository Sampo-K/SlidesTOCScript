/*

Created by Sampo Kyyrö 5.1.2022.
Yeah, I know. It's not so nice but at least it works.
If I were to do this again I would change it so that there is only one loop through all slides.
Currently the script has to loop twice: once for indexes and counts and once when adding data to TOCs.

!!!!!!!!!!!!!!!
The script should be able to handle most of changes but the most volatile part is the textbox index indicating
where is the textbox containing TOC data in the main TOC and cover pages. Currently the index is 1 in both
but it might change if templates are changed.

This information is used in two functions
1. addTitleToPage
2. getTOCSlides
Please update the targetTextBoxID value if the textbox index changes in either of these slides.
!!!!!!!!!!!!!!!

*/

function createTOC() {
  const currentPresentation = SlidesApp.getActivePresentation();
  const currentPresentationSlides = currentPresentation.getSlides();
  const TOCSlideIDs = getTOCSlides(currentPresentationSlides);

  // Add section numbering and count slides with the same title
  addIndexNumbers(currentPresentationSlides, TOCSlideIDs, 0, 1);
  // Add slide titles to cover pages and cover page titles to the main TOC.
  addTOC(currentPresentationSlides, TOCSlideIDs);
}

// Generates table of contents for both the main TOC page and cover pages.
function addTOC(currentPresentationSlides, TOCSlideIDs) {
  console.log("<==== CREATING TOC ====>");
  const mainTOCSlide = currentPresentationSlides[TOCSlideIDs[0]];
  let currentSubCoverSlide = null;

  for (slideid in currentPresentationSlides) {
    let slideID = parseInt(slideid);
    // Skip all slides before the first TOC slide
    if (slideID > TOCSlideIDs[0]) {
      
      // If the slide is a subsection cover slide
      if (TOCSlideIDs.indexOf(slideID) != -1) {
        currentSubCoverSlide = currentPresentationSlides[slideid];
        let realIndx = slideID + 1
        addTitleToPage(currentSubCoverSlide, mainTOCSlide, "MAIN");
        console.log("Current subsection cover p. " + realIndx + ": " + currentSubCoverSlide.getShapes()[0].getText().asString());
      }

      // For all other slides
      else {
        if (currentSubCoverSlide != null) {
          addTitleToPage(currentPresentationSlides[slideid], currentSubCoverSlide, "COVER");
        }
      }

    }
  }
}

// Adds a single page title to a toc page.
// Arguments: slide of the page containing the title | slide containing the toc | is this a coverpage or main toc
function addTitleToPage(slide, targetPage, type) {
    let slideTitle = getTitle(slide);
    // This allows setting different index for cover and title content textboxes.
    // Currently they are both 1.
    let targetTextBoxID = 1;
    if (type === "COVER") {
      targetTextBoxID = 1;
    }

    if (slideTitle != "" && slideTitle != null) {
      slideTitle = slideTitle.trim();
      const addedTxt = targetPage.getShapes()[targetTextBoxID].getText().appendText(slideTitle + "\n");
      // Style section headings differently than page headings
      if (type === "MAIN") {
        addedTxt.getTextStyle().setBold(true);
        addedTxt.getTextStyle().setLinkSlide(slide);
      }
      else {
        addedTxt.getTextStyle().setLinkSlide(slide);
      }
      console.log("Slide: " +  slideTitle + " added!");
    }
    else {
      console.log("[ERROR] Page title is empty!");
    }
}

// Returns the title of a given slide or null if no title can be found.
function getTitle(slide) {
  const shapes = slide.getShapes();
  let found = false;
  for (shapeid in shapes) {
    if (shapes[shapeid].getPlaceholderType() == "TITLE") {
      found = true;
      return shapes[shapeid].getText().asString().trim();
    }
  }
  if (!found) {
    console.log("[ERROR] TITLE textbox element not found!");
  }
  return null;
}

// Sets the title of a given slide.
// Arguments: slide where title will be set to | contents of the title | does the function clear all existing text or not | is this a TITLE or SUBTITLE
function setTitle(slide, newTitle, clear, type) {
  const shapes = slide.getShapes();
  let found = false;
  for (shapeid in shapes) {
    if (shapes[shapeid].getPlaceholderType() == type) {
      found = true;
      if (clear) {
        shapes[shapeid].getText().clear();
      }
      shapes[shapeid].getText().appendText(newTitle);
      console.log(shapes[shapeid].getText().asString());
    }
  }
  if (!found) {
    console.log("[ERROR] " + type + " textbox element not found!");
  }
}

// Returns a list of ids of all slides with the layout "Sisällysluettelo" or "Väliotsikko"
// Also clears existing data in all TOC pages
function getTOCSlides(allSlides) {
  let TOCSlides = [];
  for (slideid in allSlides) {
    const layoutName = allSlides[slideid].getLayout().getLayoutName();
    
    // Implemented support for different textboxes to be cleared if this is a cover page or the main toc.
    // Currently both textboxes are at index 1.
    let targetTextBoxID = 1;
    if(layoutName == "Sisällysluettelo") {
      targetTextBoxID = 1;
    }

    if (layoutName == "Sisällysluettelo" || layoutName == "Väliotsikko") {
      TOCSlides.push(parseInt(slideid));
      // Clear all TOC and cover pages from any previous content
      allSlides[slideid].getShapes()[targetTextBoxID].getText().clear();
    }
  }
  return TOCSlides;
}

// Creates numbering in the beginning of the slides in the following format: x.y. where x is the main section and y is the slide topic.
// If there are slides with the same title next to each other the script adds count in the end of the slide titles in the following format: x/y.
// where x is the number of the current slide and y amount of all slides in the set.
function addIndexNumbers(allSlides, TOCSlideIDs, mainCount, subCount) {
  console.log("<==== INDEXING ====>");
  let prevSlidesWithSameTitle = [];
  let prevTitle = "";
  let currentCoverTitle = "";
  for (slideid in allSlides) {
    let slideID = parseInt(slideid);
    // Ignore slides before the main TOC slide
    if (slideID > TOCSlideIDs[0]) {
      
      // Get the slide title and strip any previous numbering from it.
      // Please note that you should not use / in your slides as it is used to separate slide number from total amount.
      const slideTitle = removeCount(removeIndex(getTitle(allSlides[slideID])));

      // If the slide is a subsection cover slide
      if (TOCSlideIDs.indexOf(slideID) != -1) {
        mainCount += 1;
        setTitle(allSlides[slideID], mainCount + ". " + slideTitle, true, "TITLE");
        subCount = 1;
        currentCoverTitle = mainCount + ". " + slideTitle;

        if (prevSlidesWithSameTitle.length > 0) {
          prevSlidesWithSameTitle.push(allSlides[slideID-1]);
          for (slideid in prevSlidesWithSameTitle) {
            let count = parseInt(slideid) + 1;
            setTitle(prevSlidesWithSameTitle[slideid], " " + count + "/" + prevSlidesWithSameTitle.length, false, "TITLE");
          }
          prevSlidesWithSameTitle = [];
        }
      }
      
      // For all other slides
      else {
        // Check if the current title matches the previous title
        if (slideTitle == prevTitle) {
          prevSlidesWithSameTitle.push(allSlides[slideID-1]);
        }
        else if (prevSlidesWithSameTitle.length > 0) {
          prevSlidesWithSameTitle.push(allSlides[slideID-1]);
          for (slideid in prevSlidesWithSameTitle) {
            let count = parseInt(slideid) + 1;
            setTitle(prevSlidesWithSameTitle[slideid], " " + count + "/" + prevSlidesWithSameTitle.length, false, "TITLE");
          }
          prevSlidesWithSameTitle = [];
        }

        // Always add indexing to every normal slide
        setTitle(allSlides[slideID], mainCount + "." + subCount + ". " + slideTitle, true, "TITLE");
        setTitle(allSlides[slideID], currentCoverTitle, true, "SUBTITLE");
        subCount += 1;
        prevTitle = slideTitle;

        // Special case for the last slide
        if (slideID == allSlides.length - 1) {
          if (prevSlidesWithSameTitle.length > 0) {
            prevSlidesWithSameTitle.push(allSlides[slideID]);
            for (slideid in prevSlidesWithSameTitle) {
              let count = parseInt(slideid) + 1;
              setTitle(prevSlidesWithSameTitle[slideid], " " + count + "/" + prevSlidesWithSameTitle.length, false, "TITLE");
            }
            prevSlidesWithSameTitle = [];
          }
        }

      }
    }
    
  }
}

// Clears existing numbering from the beginning of the page title.
function removeIndex(title) {
  if (title != null) {
    const numIndex = title.indexOf(". ");
    if(numIndex != -1) {
      return title.substring(numIndex+1, title.length).trim();
    }
    else {
      return title;
    }
  }
  else {
    return title;
  }
}

// Clears existing counts from the end of titles of repeated slides.
function removeCount(title) {
  if (title != null) {
    const numIndex = title.indexOf("/");
    if(numIndex != -1) {
      return title.substring(0, numIndex-1).trim();
    }
    else {
      return title;
    }
  }
  else {
    return title;
  }
}
