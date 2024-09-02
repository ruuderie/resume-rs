use serde::Deserialize;
use serde_yaml;
use docx_rs::*;

const RESUME_LOCATION: &str = "/Users/ruuderie/src/git/resume-rs/Ruud_Salym_Erie_CV.yaml";

#[derive(Deserialize)]
struct Resume {
    design: Design,
    cv: CV,
    summary_of_qualifications: Vec<String>,
    experience_details: Vec<Experience>,
    technical_skills: Vec<String>,
    education_details: Vec<Education>,
}

#[derive(Deserialize)]
struct Design {
    page_size: String,
    margins: Margins,
    font: String,
    font_size: String,
}

#[derive(Deserialize)]
struct Margins {
    page: PageMargins,
}

#[derive(Deserialize)]
struct PageMargins {
    top: String,
    bottom: String,
    left: String,
    right: String,
}

#[derive(Deserialize)]
struct CV {
    personal_information: PersonalInfo,
}

#[derive(Deserialize)]
struct PersonalInfo {
    name: String,
    surname: String,
    email: String,
    phone: String,
    github: String,
    linkedin: String,
}

#[derive(Deserialize)]
struct Experience {
    company: String,
    position: String,
    location: String,
    employment_period: String,
    industry: String,
    key_responsibilities: Vec<String>,
}

#[derive(Deserialize)]
struct Education {
    institution: String,
    degree: String,
    course_name: Option<String>,
    completion_date: Option<String>,
    graduation_year: Option<String>,
    location: Option<String>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let yaml_str = std::fs::read_to_string(RESUME_LOCATION)?;
    let resume: Resume = serde_yaml::from_str(&yaml_str)?;

    let mut docx = Docx::new();

    // Apply design specifications
    apply_design(&mut docx, &resume.design);

    // Add content
    add_header(&mut docx, &resume.cv.personal_information);
    add_summary(&mut docx, &resume.summary_of_qualifications);
    add_experience(&mut docx, &resume.experience_details);
    add_skills(&mut docx, &resume.technical_skills);
    add_education(&mut docx, &resume.education_details);

    // Save the document
    let path = std::path::Path::new("resume.docx");
    let file = std::fs::File::create(path)?;
    docx.build().pack(file)?;

    println!("Resume generated successfully!");
    Ok(())
}

fn apply_design(docx: &mut Docx, design: &Design) {
    // Apply page size and margins
    let (width, height) = match design.page_size.as_str() {
        "letterpaper" => (12240, 15840), // 8.5" x 11" in twentieths of a point
        _ => (11906, 16838), // A4 default
    };
    *docx = docx.clone().page_size(width, height);
    
    let top = parse_margin(&design.margins.page.top);
    let bottom = parse_margin(&design.margins.page.bottom);
    let left = parse_margin(&design.margins.page.left);
    let right = parse_margin(&design.margins.page.right);
    *docx = docx.clone().page_margin(PageMargin::new().top(top).bottom(bottom).left(left).right(right));

    // Create styles
    let font = RunFonts::new().ascii(&design.font).hi_ansi(&design.font);
    let font_size = design.font_size.trim_end_matches("pt").parse::<u32>().unwrap_or(11) * 2;

    let normal_style = Style::new("Normal", StyleType::Paragraph)
        .name("Normal")
        .based_on("Normal")
        .fonts(font.clone())
        .size(font_size.try_into().unwrap());

    let heading_style = Style::new("Heading", StyleType::Paragraph)
        .name("Heading")
        .based_on("Normal")
        .fonts(font.clone())
        .size((font_size * 2).try_into().unwrap())
        .bold();

    let subheading_style = Style::new("Subheading", StyleType::Paragraph)
        .name("Subheading")
        .based_on("Normal")
        .fonts(font)
        .size((font_size * 3 / 2).try_into().unwrap())
        .bold();

    *docx = docx.clone()
        .add_style(normal_style)
        .add_style(heading_style)
        .add_style(subheading_style);
}

fn parse_margin(margin: &str) -> i32 {
    let value = margin.trim_end_matches(" cm").parse::<f32>().unwrap_or(2.0);
    (value * 567.0) as i32 // Convert cm to twentieths of a point
}

fn add_header(docx: &mut Docx, info: &PersonalInfo) {
    let name = Paragraph::new().add_run(
        Run::new()
            .add_text(&format!("{} {}", info.name, info.surname))
            .size(48)
            .bold()
    ).style("Heading");
    *docx = docx.clone().add_paragraph(name);

    let contact_table = Table::new(vec![
        TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&info.email))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&info.phone))),
        ]),
        TableRow::new(vec![
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&info.github))),
            TableCell::new().add_paragraph(Paragraph::new().add_run(Run::new().add_text(&info.linkedin))),
        ]),
    ]).set_grid(vec![5000, 5000]);

    *docx = docx.clone().add_table(contact_table);
    add_horizontal_line(docx);
}

fn add_summary(docx: &mut Docx, summary: &[String]) {
    add_section_title(docx, "Summary of Qualifications");

    for qual in summary {
        let para = Paragraph::new()
            .add_run(Run::new().add_text(" • "))
            .add_run(Run::new().add_text(qual))
            .style("Normal");
        *docx = docx.clone().add_paragraph(para);
    }
    add_horizontal_line(docx);
}

fn add_experience(docx: &mut Docx, experiences: &[Experience]) {
    add_section_title(docx, "WORK EXPERIENCE");

    for exp in experiences {
        let company_and_position = Paragraph::new()
            .add_run(Run::new().add_text(&exp.company).bold())
            .add_run(Run::new().add_text(" - "))
            .add_run(Run::new().add_text(&exp.position).disable_bold())
            .style("Subheading");
        *docx = docx.clone().add_paragraph(company_and_position);

        let details = Paragraph::new()
            .add_run(Run::new().add_text(&format!("{} | {}", exp.location, exp.employment_period)))
            .style("Normal");
        *docx = docx.clone().add_paragraph(details);

        let industry = Paragraph::new()
            .add_run(Run::new().add_text(&format!("Industry: {}", exp.industry)))
            .style("Normal");
        *docx = docx.clone().add_paragraph(industry);

        for resp in &exp.key_responsibilities {
            let para = Paragraph::new()
                .add_run(Run::new().add_text(" • "))  // Added two spaces before the bullet point
                .add_run(Run::new().add_text(resp))
                .style("Normal")
                .indent(Some(720), None, None, None);  // Added left indentation (720 twips = 0.5 inches)
            *docx = docx.clone().add_paragraph(para);
        }

        *docx = docx.clone().add_paragraph(Paragraph::new());
    }
    add_horizontal_line(docx);
}

fn add_skills(docx: &mut Docx, skills: &[String]) {
    add_section_title(docx, "TECHNICAL SKILLS");

    for skill in skills {
        let para = Paragraph::new()
            .add_run(Run::new().add_text("• "))
            .add_run(Run::new().add_text(skill))
            .style("Normal");
        *docx = docx.clone().add_paragraph(para);
    }
    add_horizontal_line(docx);
}

fn add_education(docx: &mut Docx, education: &[Education]) {
    add_section_title(docx, "EDUCATION");

    let mut grouped_education: std::collections::HashMap<String, Vec<&Education>> = std::collections::HashMap::new();

    for edu in education {
        grouped_education.entry(edu.institution.clone()).or_insert(Vec::new()).push(edu);
    }

    for (institution, educations) in grouped_education {
        let institution_para = Paragraph::new()
            .add_run(Run::new().add_text(&institution).bold())
            .style("Subheading");
        *docx = docx.clone().add_paragraph(institution_para);

        if let Some(location) = educations[0].location.as_ref() {
            let location_para = Paragraph::new()
                .add_run(Run::new().add_text(&format!("Location: {}", location)))
                .style("Normal");
            *docx = docx.clone().add_paragraph(location_para);
        }

        for edu in educations {
            let degree = Paragraph::new()
                .add_run(Run::new().add_text("• "))
                .add_run(Run::new().add_text(&edu.degree).italic())
                .style("Normal");
            *docx = docx.clone().add_paragraph(degree);

            if let Some(course_name) = &edu.course_name {
                let course_para = Paragraph::new()
                    .add_run(Run::new().add_text(&format!("  Course: {}", course_name)))
                    .style("Normal");
                *docx = docx.clone().add_paragraph(course_para);
            }

            if let Some(completion_date) = &edu.completion_date {
                let completion_para = Paragraph::new()
                    .add_run(Run::new().add_text(&format!("  Completed: {}", completion_date)))
                    .style("Normal");
                *docx = docx.clone().add_paragraph(completion_para);
            }

            if let Some(graduation_year) = &edu.graduation_year {
                let graduation_para = Paragraph::new()
                    .add_run(Run::new().add_text(&format!("  Graduated: {}", graduation_year)))
                    .style("Normal");
                *docx = docx.clone().add_paragraph(graduation_para);
            }
        }

        *docx = docx.clone().add_paragraph(Paragraph::new());
    }
}

fn add_section_title(docx: &mut Docx, title: &str) {
    let para = Paragraph::new()
        .add_run(Run::new().add_text(title))
        .style("Heading");
    *docx = docx.clone().add_paragraph(para);
}

fn add_horizontal_line(docx: &mut Docx) {
    let para = Paragraph::new()
        .add_run(Run::new().add_text("_").size(1))
        .align(AlignmentType::Center)
        .indent(Some(0), None, Some(0), None);
    *docx = docx.clone().add_paragraph(para);
}