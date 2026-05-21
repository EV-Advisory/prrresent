# test-create_point_slide.R

context("create_point_slide function")


# Test that a slide is created successfully
test_that("create_point_slide creates a slide successfully", {
  doc <- read_ppt_template()
  slide <- create_point_slide(doc, "Title", "Subtitle", "Body Text","", "Point - Detail Bullet", "EVA - Standard")
  expect_silent(officer:::print.rpptx(slide, tempfile(fileext = ".pptx")))
})

# Test that the function throws an error when the layout doesn't exist
test_that("create_point_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  expect_error(create_point_slide(doc, "Title", "Subtitle", "Body Text", "","Non-existent layout", "EVA - Standard"))
})

# Test that the function throws an error when the layout is not a point slide
test_that("create_point_slide throws an error when the layout is not a point slide", {
  doc <- read_ppt_template()
  expect_error(create_point_slide(doc, "Title", "Subtitle", "Body Text","", "Title - Black", "EVA - Standard"))
})

# Add more test cases as needed
