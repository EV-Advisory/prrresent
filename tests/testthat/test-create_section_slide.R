# test-create_section_slide.R

context("create_section_slide function")

# Test that a section slide is created successfully
test_that("create_section_slide creates a slide successfully", {
  doc <- read_ppt_template()
  slide <- create_section_slide(
    doc,
    slide_title = "Section Title",
    slide_subtitle = "Section Subtitle"
  )
  expect_silent(officer:::print.rpptx(slide, tempfile(fileext = ".pptx")))
})

# Test that the function throws an error when the layout doesn't exist
test_that("create_section_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  expect_error(create_section_slide(doc, layout_name = "Non-existent layout",
                                    slide_title = "Section Title", slide_subtitle = "Section Subtitle"))
})

# Test that the function throws an error when the layout is not a section slide
test_that("create_section_slide throws an error when the layout is not a section slide", {
  doc <- read_ppt_template()
  expect_error(create_section_slide(doc, layout_name = "Title - Black",
                                    slide_title = "Section Title", slide_subtitle = "Section Subtitle"))
})

# Add more test cases as needed
