# test-create_graphic_slide
context("create_graphic_slide function")
require(ggplot2)
# test that a slide is created successfully
test_that("create_graphic_slide creates a slide successfully", {
  doc <- read_ppt_template()
  slide_graphic <- qplot(x = wt, y = mpg, data = mtcars)
  slide <- create_graphic_slide(x = doc,
                                slide_title = "Test Slide",
                                slide_caption = "This is a test slide",
                                slide_graphic = slide_graphic,
                                slide_body = "This is the body text",
                                slide_section = "Test Section",
                                footer_text = "Footer text",
                                layout_name = "Graphic Full",
                                master_name = 'EVA - Standard')
  expect_silent(officer:::print.rpptx(slide,tempfile(fileext = ".pptx")))
})

# test that the function throws an error when the layout doesn't exist
test_that("create_graphic_slide throws an error when the layout doesn't exist", {
  doc <- read_ppt_template()
  slide_graphic <-qplot(x = wt, y = mpg, data = mtcars)
  expect_error(create_graphic_slide(x = doc,
                                    slide_title = "Test Slide",
                                    slide_caption = "This is a test slide",
                                    slide_graphic = slide_graphic,
                                    slide_body = "This is the body text",
                                    slide_section = "Test Section",
                                    footer_text = "Footer text",
                                    layout_name = "Non-existent layout",
                                    master_name = 'EVA - Standard'))
})

# test that the function throws an error when the layout is not a graphic slide
test_that("create_graphic_slide throws an error when the layout is not a graphic slide", {
  doc <- read_ppt_template()
  slide_graphic <- qplot(x = wt, y = mpg, data = mtcars)
  expect_error(create_graphic_slide(x = doc,
                                    slide_title = "Test Slide",
                                    slide_caption = "This is a test slide",
                                    slide_graphic = slide_graphic,
                                    slide_body = "This is the body text",
                                    slide_section = "Test Section",
                                    footer_text = "Footer text",
                                    layout_name = "Title - Black",
                                    master_name = 'EVA - Standard'))
})
