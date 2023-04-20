# test-create_title_slide.R

context("create_title_slide function")

# Test that a slide is created successfully
test_that("create_title_slide creates a slide successfully", {
  doc <- read_ppt_template()
  slide <- create_title_slide(doc, "Title", "Subtitle", system.file("extdata/cat.jpeg",package = "prrresent"), "Title - Black", "EVA - Standard")
  expect_silent(officer:::print.rpptx(slide, tempfile(fileext = ".pptx")))
})

# Test that the function throws an error when the layout doesn't exist
test_that("create_title_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  expect_error(create_title_slide(doc, "Title", "Subtitle", system.file("extdata/cat.jpeg",package = "prrresent"), "Non-existent layout", "EVA - Standard"))
})

# Test that the function throws an error when the layout is not a title slide
test_that("create_title_slide throws an error when the layout is not a title slide", {
  doc <- read_ppt_template()
  expect_error(create_title_slide(doc, "Title", "Subtitle", system.file("extdata/cat.jpeg",package = "prrresent"), "Point - Detail Bullet", "EVA - Standard"))
})

