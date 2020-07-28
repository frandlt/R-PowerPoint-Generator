library(officer)

# Crates a new PowerPoint file
pres <- read_pptx()

# Slide N°1 ---------------------------------------------------------------

# Add Title Only slide
pres <- add_slide(pres, layout = "Title Only", master = "Office Theme")

# Add Title text
pres <- ph_with(pres, value = "My first presentation", location = ph_location_type(type = "title"))


# Slide N°2 ---------------------------------------------------------------

# Add Title and Content slide
pres <- add_slide(pres, layout = "Title and Content", master = "Office Theme")

# Add Title and Content text
pres <- ph_with(pres, value = "This is the second slide", location = ph_location_type(type = "title"))
pres <- ph_with(pres, value = c("First line", "Second Line", "Third line"), location = ph_location_type(type = "body"))


# Slide N°3 ---------------------------------------------------------------

# Create sample data frame
frame <- data.frame(a = 1:5, b = 11:15, c = 21:25)

# Create slide to hold table
pres <- add_slide(pres, layout = "Title and Content", master = "Office Theme")
pres <- ph_with(pres, value = "Table Example", location = ph_location_type(type = "title"))

# Add dataframe to PowerPoint slide
pres <- ph_with(pres, value = frame, location = ph_location_type(type = "body"))


# Generate PowerPoint file ------------------------------------------------

print(pres, target = "example.pptx")

# End of Script