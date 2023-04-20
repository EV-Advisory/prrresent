# test-create_agenda_slide

context("create_agenda_slide function")

# Test that a slide is created successfully
test_that("create_agenda_slide creates a slide successfully", {
  doc <- read_ppt_template()
  slide <- create_agenda_slide(doc, "Agenda", section_title = "Introduction")
  expect_silent(officer:::print.rpptx(slide, tempfile(fileext = ".pptx")))
})

# Test that the function throws an error when the layout doesn't exist
test_that("create_agenda_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  expect_error(create_agenda_slide(doc, "Agenda", layout_name = "Non-existent layout"))
})

# Test that the function throws an error when the layout is not an agenda slide
test_that("create_agenda_slide throws an error when the layout is not an agenda slide", {
  doc <- read_ppt_template()
  expect_error(create_agenda_slide(doc, "Agenda", layout_name = "Title - Black"))
})

# Add more test cases as needed
