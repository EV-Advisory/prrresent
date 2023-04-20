# test-create_column_slide.R

context("create_column_slide function")

# Test that a slide is created successfully
test_that("create_column_slide creates a slide successfully", {
  doc <- read_ppt_template()
  slide <- create_column_slide(
    doc,
    column_data = list(
      list('title' = "Title of Column 1", 'body' = "Body of Column 1"),
      list('title' = "Title of Column 2", 'body' = "Body of Column 2")
    ),
    layout_name = "Two Column"
  )
  expect_silent(officer:::print.rpptx(slide, tempfile(fileext = ".pptx")))
})

# Test that the function throws an error when the layout doesn't exist
test_that("create_column_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  expect_error(create_column_slide(doc, layout_name = "Non-existent layout"))
})

# Test that the function throws an error when the layout is not a column slide
test_that("create_column_slide throws an error when the layout is not a column slide", {
  doc <- read_ppt_template()
  expect_error(create_column_slide(doc, layout_name = "Title - Black"))
})

# Test that the function throws an error when column_data does not match layout
test_that("create_column_slide throws an error when column_data does not match layout", {
  doc <- read_ppt_template()
  expect_error(create_column_slide(
    doc,
    column_data = list(
      list('title' = "Title of Column 1", 'body' = "Body of Column 1"),
      list('title' = "Title of Column 2", 'body' = "Body of Column 2"),
      list('title' = "Title of Column 3", 'body' = "Body of Column 3")
    ),
    layout_name = "Two Column"
  ))
})

