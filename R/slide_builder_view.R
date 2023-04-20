#' Slide Builder View Function
#'
#' This function creates a visual representation of slide layout properties for a PowerPoint slide
#' using ggplot2. The function takes the layout properties of the slide and creates a ggplot
#' object with rectangles representing the different placeholders and labels.
#'
#' @param layout_properties A data frame containing the layout properties of a PowerPoint slide.
#' This can be obtained using the `layout_properties()` function from the officer package.
#' @return A ggplot object displaying the slide layout with rectangles and labels for the placeholders.
#' @importFrom ggplot2 ggplot aes geom_rect geom_text theme_void
#' @importFrom officer layout_properties
#' @examples
#' \dontrun{
#' # Assuming 'doc' is an officer presentation object
#' slide_to_build <- layout_properties(doc, "Two Column - No Title")
#' slide_builder_view(slide_to_build)
#' }
#' @export
slide_builder_view <-
  function(layout_properties) {
    ggplot(layout_properties,
           aes(
             xmin = offx,
             ymin = -offy,
             xmax = offx + cx,
             ymax = -offy - cy
           )) +
      geom_rect(fill = "pink") +
      geom_text(
        aes(
          x = offx,
          y = -offy - cy / 2,
          label = ph_label
        ),
        color = "black",
        size = 3.5,
        hjust = 0
      ) +
      theme_void()
  }
