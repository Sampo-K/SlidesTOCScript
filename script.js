function createTOC() {
  const currentPresentation = SlidesApp.getActivePresentation();
  const currentPresentationSlides = currentPresentation.getSlides();
  const TOCSlideIDs = getTOCSlides(currentPresentationSlides);

  // Add section numbering and count slides with the same title
  addIndexNumbers(currentPresentationSlides, TOCSlideIDs, 0, 1);
  // Add slide titles to cover pages and cover page titles to the main TOC.
  addTOC(currentPresentationSlides, TOCSlideIDs);
}

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

function addTitleToPage(slide, targetPage, type) {
    let slideTitle = getTitle(slide);
    let targetTextBoxID = 1;
    // This allows setting different index for cover and title content textboxes.
    // Currently they are both 1.
    if (type === "COVER") {
      targetTextBoxID = 1;
    }

    if (slideTitle != "" && slideTitle != null) {
      slideTitle = slideTitle.trim();
      const addedTxt = targetPage.getShapes()[targetTextBoxID].getText().appendText("\n" + slideTitle);
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

function getTitle(slide) {
  const shapes = slide.getShapes();
  for (shapeid in shapes) {
    if (shapes[shapeid].getPlaceholderType() == "TITLE") {
      return shapes[shapeid].getText().asString().trim();
    }
    else {
      //console.log("[ERROR] TITLE textbox element not found!");
    }
  }
  return null;
}

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

function getTOCSlides(allSlides) {
  let TOCSlides = [];
  for (slideid in allSlides) {
    const layoutName = allSlides[slideid].getLayout().getLayoutName();
    if (layoutName == "Sisällysluettelo" || layoutName == "Väliotsikko") {
      TOCSlides.push(parseInt(slideid));
      // Clear all TOC and cover pages from any previous content
      allSlides[slideid].getShapes()[1].getText().clear();
    }
  }
  return TOCSlides;
}

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
