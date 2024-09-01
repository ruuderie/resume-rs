use serde::Deserialize;
use serde_yaml;
use std::fs::File;
use std::io::Write;

const RESUME_LOCATION: &str = "/Users/ruuderie/src/git/resume-rs/Ruud_Salym_Erie_CV.yaml";

#[derive(Deserialize, Debug)]
struct Resume {
    cv: CV,
}

#[derive(Deserialize, Debug)]
struct CV {
    personal_information: PersonalInfo,
    #[serde(default)]
    summary_of_qualifications: Option<Vec<String>>,
    experience_details: Vec<Experience>,
    #[serde(default)]
    projects: Vec<Project>,
    technical_skills: Vec<String>,
    education_details: Vec<Education>,
}

#[derive(Deserialize, Debug)]
struct PersonalInfo {
    name: String,
    surname: String,
    github: String,
}

#[derive(Deserialize, Debug)]
struct Experience {
    position: String,
    company: String,
    employment_period: String,
    location: String,
    industry: String,
    key_responsibilities: Vec<String>,
}

#[derive(Deserialize, Debug)]
struct Project {
    name: String,
    details: Vec<String>,
}

#[derive(Deserialize, Debug)]
struct Education {
    degree: String,
    institution: String,
    #[serde(default)]
    graduation_year: Option<String>,
    #[serde(default)]
    location: Option<String>,
    #[serde(default)]
    course_name: Option<String>,
    #[serde(default)]
    completion_date: Option<String>,
    #[serde(default)]
    instructor: Option<String>,
    #[serde(default)]
    credential_id: Option<String>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let yaml_str = std::fs::read_to_string(RESUME_LOCATION)?;
    let resume: Resume = serde_yaml::from_str(&yaml_str)?;

    let mut output = String::new();

    // Header
    output.push_str(&format!(
        "# {} {}\n",
        resume.cv.personal_information.name, resume.cv.personal_information.surname
    ));
    output.push_str(&format!("{}\n\n", resume.cv.personal_information.github));

    // Summary of Qualifications (if present)
    if let Some(summary) = &resume.cv.summary_of_qualifications {
        output.push_str("## Summary of Qualifications\n\n");
        for qual in summary {
            output.push_str(&format!("- {}\n", qual));
        }
        output.push_str("\n");
    }

    // Experience
    output.push_str("## Experience\n\n");
    for exp in &resume.cv.experience_details {
        output.push_str(&format!("### {}\n", exp.company));
        output.push_str(&format!("**{}**\n", exp.position));
        output.push_str(&format!("{} | {}\n", exp.location, exp.employment_period));
        output.push_str(&format!("Industry: {}\n\n", exp.industry));
        for resp in &exp.key_responsibilities {
            output.push_str(&format!("- {}\n", resp));
        }
        output.push_str("\n");
    }

    // Projects (if present)
    if !resume.cv.projects.is_empty() {
        output.push_str("## Projects\n\n");
        for project in &resume.cv.projects {
            output.push_str(&format!("### {}\n", project.name));
            for detail in &project.details {
                output.push_str(&format!("- {}\n", detail));
            }
            output.push_str("\n");
        }
    }

    // Technical Skills
    output.push_str("## Technical Skills\n\n");
    for skill in &resume.cv.technical_skills {
        output.push_str(&format!("- {}\n", skill));
    }
    output.push_str("\n");

    // Education
    output.push_str("## Education\n\n");
    for edu in &resume.cv.education_details {
        output.push_str(&format!("### {}\n", edu.institution));
        output.push_str(&format!("**{}**\n", edu.degree));
        if let Some(course_name) = &edu.course_name {
            output.push_str(&format!("Course: {}\n", course_name));
        }
        if let Some(instructor) = &edu.instructor {
            output.push_str(&format!("Instructor: {}\n", instructor));
        }
        if let Some(completion_date) = &edu.completion_date {
            output.push_str(&format!("Completed: {}\n", completion_date));
        }
        if let Some(graduation_year) = &edu.graduation_year {
            output.push_str(&format!("Graduated: {}\n", graduation_year));
        }
        if let Some(location) = &edu.location {
            output.push_str(&format!("Location: {}\n", location));
        }
        if let Some(credential_id) = &edu.credential_id {
            output.push_str(&format!("Credential ID: {}\n", credential_id));
        }
        output.push_str("\n");
    }

    // Write to file
    let mut file = File::create("resume.md")?;
    file.write_all(output.as_bytes())?;

    println!("Resume generated successfully!");
    Ok(())
}
