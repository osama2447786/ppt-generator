import streamlit as st
from pptx import Presentation
import io

def create_ppt(title, slides_data):
    # Initialize presentation
    prs = Presentation()
    
    # Create Title Slide (Layout 0)
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title_shape = slide.shapes.title
    subtitle = slide.placeholders[1]
    title_shape.text = title
    subtitle.text = "Generated with Streamlit"
    
    # Create Content Slides (Layout 1)
    for slide_title, content in slides_data:
        bullet_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes
        
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        
        title_shape.text = slide_title
        tf = body_shape.text_frame
        tf.text = content
        
    # Save the presentation to an in-memory buffer
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    
    return ppt_stream

# --- Streamlit User Interface ---
st.title("📊 AI/Backend PPT Generator")
st.write("Add your slide details below to instantly generate a PowerPoint file.")

# Main Title
deck_title = st.text_input("Presentation Title", "Quarterly Update")

# Dynamic Slide Inputs
st.markdown("### Slide Content")
num_slides = st.number_input("How many content slides do you want?", min_value=1, max_value=10, value=2)

slides_data = []
for i in range(int(num_slides)):
    with st.expander(f"Slide {i+1} Setup", expanded=True):
        slide_title = st.text_input("Slide Title", value=f"Slide {i+1}", key=f"title_{i}")
        slide_content = st.text_area("Slide Content", value="Type your points here...", key=f"content_{i}")
        slides_data.append((slide_title, slide_content))

# Generate and Download
if st.button("Generate PowerPoint"):
    with st.spinner("Generating deck..."):
        ppt_file = create_ppt(deck_title, slides_data)
        
        st.success("Your presentation is ready!")
        st.download_button(
            label="⬇️ Download .pptx File",
            data=ppt_file,
            file_name="generated_deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
