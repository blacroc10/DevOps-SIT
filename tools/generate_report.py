from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
SCREENSHOTS = ROOT / "screenshots"
OUTPUT = ROOT / "DevOps Lab Record - Shubhankar Sarangi.docx"
LOGO = ROOT / "assets" / "sit-logo.png"

STUDENT_NAME = "Shubhankar Sarangi"
PRN = "23070122206"
FACULTY = "Prof. Arijit Dutta"
REPOSITORY = "https://github.com/blacroc10/DevOps-SIT"


ASSIGNMENTS = [
    {
        "number": 1,
        "title": "Git Setup & Basic Workflow",
        "date": "16th July, 2026",
        "theory": [
            "Git is a distributed version control system that records source-code changes as commits and allows every clone to retain the complete project history.",
            "The working tree contains edited files, the staging area selects the next snapshot, and the local repository stores committed snapshots.",
            "GitHub hosts the remote repository so that committed work can be synchronized and reviewed from any authorized machine.",
            "A clear commit message and a clean working tree make project history easier to audit and reproduce.",
        ],
        "commands": [
            "git init",
            "git config user.name \"Shubhankar Sarangi\"",
            "git add README.md .gitignore",
            "git commit -m \"Add initial README and repository configuration\"",
            "git remote add origin https://github.com/blacroc10/DevOps-SIT.git",
            "git push -u origin main",
            "git log --oneline",
        ],
        "steps": [
            ("01-repository-setup.png", "Initialize the repository and configure the remote", "The local repository, identity, initial files, and GitHub remote were configured."),
            ("02-commit-verification.png", "Commit and verify the basic Git workflow", "The files were staged and committed, and the clean working tree and commit history were verified."),
            ("03-github-commit-history.png", "Verify commits on GitHub", "The GitHub commit history confirms that the repository was pushed successfully under the student's account."),
        ],
    },
    {
        "number": 2,
        "title": "Git Branching & Merging",
        "date": "16th July, 2026",
        "theory": [
            "A Git branch is a movable pointer to a commit and provides an isolated line of development.",
            "Feature branches allow changes to be developed and reviewed without destabilizing the main branch.",
            "A merge combines the history and content of another branch into the currently checked-out branch.",
            "Deleting a merged feature branch removes the temporary pointer without deleting its merged commits.",
        ],
        "commands": [
            "git switch -c feature-update",
            "git add README.md",
            "git commit -m \"Update README on feature branch\"",
            "git switch main",
            "git merge feature-update",
            "git branch -d feature-update",
            "git log --oneline --graph --all",
        ],
        "steps": [
            ("01-create-feature-branch.png", "Create and work on a feature branch", "A separate feature branch was created and used for the README update."),
            ("02-commit-and-merge.png", "Commit the feature and merge it into main", "The feature commit was merged into main and the resulting branch history was verified."),
        ],
    },
    {
        "number": 3,
        "title": "Simulating Git Collaboration & Conflict Resolution",
        "date": "16th July, 2026",
        "theory": [
            "A merge conflict occurs when Git cannot safely combine competing edits to the same region of a file.",
            "Conflict markers identify the current branch content, the separator, and the incoming branch content.",
            "The developer must choose or combine the intended text, remove all conflict markers, stage the resolved file, and complete the merge commit.",
            "Reviewing the final file and graph confirms that no edits were lost and both lines of development are represented.",
        ],
        "commands": [
            "git switch -c conflicting-feature",
            "git switch main",
            "git merge conflicting-feature",
            "git status",
            "git add README.md",
            "git commit -m \"Resolve README merge conflict\"",
            "git log --oneline --graph --all",
        ],
        "steps": [
            ("01-main-change.png", "Create a competing change on main", "The same README line was edited differently to simulate concurrent collaboration."),
            ("02-merge-conflict-01.png", "Attempt the merge and observe the conflict", "Git stopped the merge because the two branches changed the same content."),
            ("02-merge-conflict-02.png", "Inspect the conflict markers", "The current and incoming versions were reviewed before resolving the file."),
            ("03-conflict-resolution.png", "Resolve, commit, and verify the merge", "The final text was selected, staged, committed, and confirmed in the branch graph."),
        ],
    },
    {
        "number": 4,
        "title": "Jira Project & Issue Tracking",
        "date": "16th July, 2026",
        "theory": [
            "Jira is a work-management platform used to plan, track, prioritize, and report software tasks.",
            "A Task represents a unit of work, a Story captures user-facing value, and a Bug records incorrect behavior.",
            "A Kanban board visualizes work through workflow columns such as To Do, In Progress, and Done.",
            "Moving work items between statuses provides a transparent history of progress and helps identify bottlenecks.",
        ],
        "commands": [
            "Create project: DevOps Lab Project (key: KAN)",
            "Create Task: Set up database",
            "Create Story: Develop login page",
            "Create Bug: Submit button not working",
            "Move each work item: To Do -> In Progress -> Done",
        ],
        "steps": [
            ("01-board-created-items.png", "Create Task, Story, and Bug work items", "The board shows all three requested work-item types in their initial workflow states."),
            ("02-board-in-progress.png", "Move the work items to In Progress", "All three work items were moved to the active work column."),
            ("03-board-done.png", "Complete the workflow", "The Task, Story, and Bug were moved to Done, completing the Kanban workflow."),
        ],
    },
    {
        "number": 5,
        "title": "Continuous Integration with Jenkins",
        "date": "16th July, 2026",
        "theory": [
            "Continuous Integration automatically builds and tests changes frequently so integration problems are discovered early.",
            "Jenkins is an automation server that executes jobs on a controller or configured agents and records console output for each build.",
            "A Freestyle project provides a graphical job configuration for source control, triggers, environment settings, and build steps.",
            "A successful console log and green build status provide traceable evidence that the configured steps completed.",
        ],
        "commands": [
            "docker build -t devops-jenkins:lab jenkins",
            "docker run -d --name devops-jenkins -p 8090:8080 devops-jenkins:lab",
            "Create Freestyle project: devops-freestyle",
            "Execute shell step: echo \"Hello from Jenkins CI\"",
            "Build Now and inspect Console Output",
        ],
        "steps": [
            ("01-jenkins-container.png", "Start the Jenkins controller container", "A dedicated Jenkins controller image and container were prepared for the lab."),
            ("02-jenkins-ready.png", "Verify the Jenkins dashboard", "Jenkins loaded successfully on localhost and was ready to run jobs."),
            ("03-freestyle-build-log-01.png", "Run the Freestyle job", "The configured shell step started and produced console output."),
            ("03-freestyle-build-log-02.png", "Inspect the completed console log", "The build log shows the command output and successful completion."),
            ("04-freestyle-success.png", "Confirm the successful build", "The job dashboard records the build as successful."),
        ],
    },
    {
        "number": 6,
        "title": "Basic Declarative Pipeline",
        "date": "16th July, 2026",
        "theory": [
            "Pipeline as Code stores the CI process in version control so it can be reviewed and reproduced.",
            "A declarative Jenkins pipeline uses a pipeline block containing an agent, stages, and individual steps.",
            "Stages divide the workflow into understandable phases such as Build, Test, and Deploy.",
            "Jenkins records each stage independently, making failures and execution duration easy to identify.",
        ],
        "commands": [
            "pipeline { agent any; stages { ... } }",
            "stage('Build') { steps { echo 'Building...' } }",
            "stage('Test') { steps { echo 'Testing...' } }",
            "stage('Deploy') { steps { echo 'Deploying...' } }",
            "Run the pipeline and inspect Stage View",
        ],
        "steps": [
            ("01-declarative-console-01.png", "Execute the declarative pipeline", "Jenkins parsed the pipeline and began executing the declared stages."),
            ("01-declarative-console-02.png", "Verify stage output in the console", "The build, test, and deployment messages completed without errors."),
            ("02-pipeline-stages.png", "Review the pipeline stage view", "The visual stage view confirms that every stage completed successfully."),
        ],
    },
    {
        "number": 7,
        "title": "Pipeline with Parameters",
        "date": "23rd July, 2026",
        "theory": [
            "Pipeline parameters allow a single Jenkins job to accept user input at build time.",
            "String, choice, Boolean, and password parameters can control environments, versions, or optional behavior.",
            "Declarative pipelines define parameters in a parameters block and access values through the params object.",
            "Parameterized builds improve reuse because the pipeline logic does not need to be edited for every execution.",
        ],
        "commands": [
            "parameters { string(name: 'GREETING_NAME', defaultValue: 'Shubhankar') }",
            "echo \"Hello ${params.GREETING_NAME}\"",
            "Select Build with Parameters",
            "Enter a parameter value and start the build",
        ],
        "steps": [
            ("01-parameterized-console-01.png", "Configure and start a parameterized build", "The build was started with an explicit parameter value."),
            ("01-parameterized-console-02.png", "Execute the parameter-aware pipeline", "Jenkins evaluated the supplied value while running the pipeline."),
            ("02-parameter-output.png", "Verify parameter output", "The console output confirms that the pipeline used the selected parameter."),
        ],
    },
    {
        "number": 8,
        "title": "Dockerfile & Image Build",
        "date": "23rd July, 2026",
        "theory": [
            "A Dockerfile is a version-controlled recipe for building an immutable container image.",
            "Instructions such as FROM, WORKDIR, COPY, RUN, EXPOSE, and CMD create ordered image layers.",
            "An image packages the application and dependencies, while a container is a running instance of that image.",
            "Port publishing maps a host port to the application port inside the container.",
        ],
        "commands": [
            "docker build -t devops-sit-flask:1.0 my-flask-app",
            "docker image ls devops-sit-flask",
            "docker run -d --name flask-lab -p 8081:5000 devops-sit-flask:1.0",
            "docker ps",
            "Open http://localhost:8081",
        ],
        "steps": [
            ("01-image-build.png", "Build and run the Flask image", "Docker built the application image and started a container from it."),
            ("02-flask-browser.png", "Verify the containerized application", "The Flask response was reached through the published host port."),
        ],
    },
    {
        "number": 9,
        "title": "Container Management & Networking",
        "date": "23rd July, 2026",
        "theory": [
            "Docker manages containers through a lifecycle of create, start, inspect, stop, and remove operations.",
            "A user-defined bridge network provides automatic DNS resolution between containers by container name.",
            "Network isolation limits communication to attached containers while published ports expose selected services to the host.",
            "Inspection and connectivity tests verify addresses, attached containers, and application reachability.",
        ],
        "commands": [
            "docker network create devops-network",
            "docker run -d --name flask-networked --network devops-network -p 8083:5000 devops-sit-flask:1.0",
            "docker network inspect devops-network",
            "docker exec flask-networked hostname",
            "docker ps",
        ],
        "steps": [
            ("01-container-networking.png", "Create and inspect the Docker network", "The application container was attached to a dedicated user-defined bridge network."),
            ("02-browser-access.png", "Verify the networked container", "The published service remained reachable from the host browser."),
        ],
    },
    {
        "number": 10,
        "title": "Docker Volumes & Data Persistence",
        "date": "23rd July, 2026",
        "theory": [
            "Container writable layers are ephemeral and are removed with the container.",
            "A Docker volume stores data independently of a container and can be mounted by replacement containers.",
            "Named volumes are managed by Docker and are suitable for application state such as databases and uploaded files.",
            "Persistence is verified by writing data, removing the original container, and reading the same data from a new container.",
        ],
        "commands": [
            "docker volume create devops-data",
            "docker run --rm -v devops-data:/data alpine sh -c \"echo persistent-data > /data/message.txt\"",
            "docker run --rm -v devops-data:/data alpine cat /data/message.txt",
            "docker volume inspect devops-data",
        ],
        "steps": [
            ("01-volume-persistence.png", "Create and test a named volume", "Data was written to the volume and recovered from a separate container."),
            ("02-volume-browser.png", "Verify the persistent application result", "The application workflow demonstrated that state remained available independently of the container."),
        ],
    },
    {
        "number": 11,
        "title": "Docker Compose for Multi-Container App",
        "date": "23rd July, 2026",
        "theory": [
            "Docker Compose defines a multi-container application in one YAML file.",
            "Services describe images, builds, ports, environment variables, dependencies, networks, and volumes.",
            "Compose creates a shared default network so the Flask service can resolve the Redis service by name.",
            "A named Redis volume preserves data while docker compose up and down control the complete application stack.",
        ],
        "commands": [
            "docker compose -f compose-app/compose.yaml up -d --build",
            "docker compose -f compose-app/compose.yaml ps",
            "docker compose -f compose-app/compose.yaml logs",
            "Open http://localhost:8082",
            "docker compose -f compose-app/compose.yaml down",
        ],
        "steps": [
            ("01-compose-services-01.png", "Build the Compose services", "Compose built the Flask service and prepared the Redis dependency."),
            ("01-compose-services-02.png", "Start and inspect the service stack", "Both services reached the running state on the shared Compose network."),
            ("02-compose-browser.png", "Verify Flask and Redis integration", "The browser response confirms successful communication between the web and Redis containers."),
        ],
    },
    {
        "number": 12,
        "title": "Kubernetes Pod & Deployment",
        "date": "23rd July, 2026",
        "theory": [
            "A Pod is Kubernetes' smallest deployable unit and contains one or more tightly coupled containers.",
            "A Deployment manages a desired number of identical Pod replicas through a ReplicaSet.",
            "Labels and selectors associate the Deployment with its Pods, while readiness probes control traffic eligibility.",
            "Declarative manifests make the desired cluster state reviewable, repeatable, and self-healing.",
        ],
        "commands": [
            "kind create cluster --name devops-lab",
            "kind load docker-image devops-sit-flask:1.0 --name devops-lab",
            "kubectl apply -f kubernetes/deployment.yaml",
            "kubectl get deployments,pods -o wide",
            "kubectl describe deployment flask-app-deployment",
        ],
        "steps": [
            ("01-kubernetes-deployment-01.png", "Create the cluster and apply the Deployment", "The application image was loaded into Kind and the Deployment manifest was applied."),
            ("01-kubernetes-deployment-02.png", "Verify Deployment replicas and Pods", "Kubernetes reported the desired replicas as available and ready."),
        ],
    },
    {
        "number": 13,
        "title": "Kubernetes Service & Networking",
        "date": "23rd July, 2026",
        "theory": [
            "Pods are replaceable, so a Kubernetes Service provides a stable virtual IP and DNS name.",
            "A Service selector discovers matching Pods and load-balances traffic across their endpoints.",
            "port is the Service-facing port, targetPort is the container port, and NodePort exposes a fixed port on each node.",
            "Endpoints and an HTTP request verify that the Service routes traffic to ready application Pods.",
        ],
        "commands": [
            "kubectl apply -f kubernetes/service.yaml",
            "kubectl get service flask-app-service",
            "kubectl get endpoints flask-app-service",
            "kubectl port-forward service/flask-app-service 8084:80",
            "Open http://localhost:8084",
        ],
        "steps": [
            ("01-kubernetes-service.png", "Create and inspect the Kubernetes Service", "The NodePort Service selected the Flask Pods and exposed port 80 to the cluster."),
            ("02-service-browser.png", "Verify Service routing", "A browser request routed through the Service and reached the deployed Flask application."),
        ],
    },
    {
        "number": 14,
        "title": "CI/CD: Jenkins Pipeline Building a Docker Image",
        "date": "23rd July, 2026",
        "theory": [
            "CI/CD combines automated integration checks with repeatable delivery steps.",
            "The Jenkins controller image includes the Docker CLI and uses the host Docker socket to invoke the Docker engine.",
            "The version-controlled Jenkinsfile checks out the repository, builds a uniquely tagged image, and validates it.",
            "A successful pipeline links a source revision, Jenkins build number, and produced Docker image for traceability.",
        ],
        "commands": [
            "checkout scm",
            "docker build -t devops-sit-flask:${BUILD_NUMBER} my-flask-app",
            "docker image inspect devops-sit-flask:${BUILD_NUMBER}",
            "Run the Jenkins pipeline from the repository Jenkinsfile",
        ],
        "steps": [
            ("01-jenkins-docker-build.png", "Build the Docker image from Jenkins", "The pipeline checked out the repository and executed the Docker image build."),
            ("02-pipeline-stages.png", "Verify CI/CD stages", "The Jenkins stage view confirms successful checkout, image build, and validation."),
        ],
    },
    {
        "number": 15,
        "title": "Basic Infrastructure as Code with Terraform",
        "date": "23rd July, 2026",
        "theory": [
            "Infrastructure as Code describes resources in version-controlled configuration instead of manual procedures.",
            "Terraform uses providers to manage resources and maintains state to compare configuration with real infrastructure.",
            "terraform init installs providers, plan previews changes, apply converges the state, and destroy removes managed resources.",
            "An update plan demonstrates idempotent change management, while the final destroy proves lifecycle cleanup.",
        ],
        "commands": [
            "terraform -chdir=terraform init",
            "terraform -chdir=terraform fmt -check",
            "terraform -chdir=terraform validate",
            "terraform -chdir=terraform plan",
            "terraform -chdir=terraform apply -auto-approve",
            "terraform -chdir=terraform destroy -auto-approve",
        ],
        "steps": [
            ("01-terraform-workflow-01.png", "Initialize and validate Terraform", "Terraform installed the required local provider and validated the configuration."),
            ("01-terraform-workflow-02.png", "Review the execution plan", "The plan described the local file resource before any change was applied."),
            ("01-terraform-workflow-03.png", "Apply and verify the resource", "Terraform created the managed file and displayed the configured output."),
            ("01-terraform-workflow-04.png", "Update the managed infrastructure", "A changed configuration produced an update plan and was applied through Terraform."),
            ("01-terraform-workflow-05.png", "Destroy and verify cleanup", "Terraform removed the managed resource and the final filesystem check confirmed its deletion."),
        ],
    },
]


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=90, start=100, bottom=90, end=100):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_width(cell, inches):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(int(inches * 1440)))
    tc_w.set(qn("w:type"), "dxa")


def set_run_font(run, size=10.5, bold=None, italic=None, color=None, name="Times New Roman"):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)


def format_paragraph(paragraph, before=0, after=3, line=1.05):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def add_text(
    paragraph,
    text,
    size=10.5,
    bold=False,
    italic=False,
    name="Times New Roman",
    color=None,
):
    run = paragraph.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic, name=name, color=color)
    return run


def clear_cell(cell):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    format_paragraph(paragraph, after=0)
    return paragraph


def add_header(document, number, logo_path):
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p, after=0, line=1)
    p.add_run().add_picture(str(logo_path), width=Inches(1.13))

    header_lines = [
        ("SYMBIOSIS INSTITUTE OF TECHNOLOGY, PUNE", 16, True),
        ("Symbiosis International (Deemed University)", 13.5, True),
        ("(Established under section 3 of the UGC Act, 1956)", 8, False),
        ("Re-accredited by NAAC with 'A' grade (3.58/4) | Awarded Category - I by UGC", 7, True),
        ("Founder: Prof. Dr. S. B. Mujumdar, M. Sc., Ph. D. (Awarded Padma Bhushan and Padma Shri by President of India)", 7, True),
    ]
    for text, size, bold in header_lines:
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p, after=0, line=1)
        add_text(p, text, size=size, bold=bold)

    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(p, before=3, after=4, line=1)
    add_text(p, f"Assignment No. {number:02d}", size=14, bold=True)


def add_label(cell, text):
    p = clear_cell(cell)
    add_text(p, text, size=11, bold=True)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_shading(cell, "EDEDED")


def add_value(cell, text, bold=False):
    p = clear_cell(cell)
    add_text(p, text, size=11, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_theory(cell, assignment):
    clear_cell(cell)
    heading = cell.paragraphs[0]
    add_text(heading, "Key concepts", size=11, bold=True, color=(128, 0, 0))
    format_paragraph(heading, after=3)
    for item in assignment["theory"]:
        p = cell.add_paragraph()
        format_paragraph(p, after=2, line=1.08)
        add_text(p, f"• {item}", size=10.5)
    p = cell.add_paragraph()
    format_paragraph(p, before=3, after=2)
    add_text(p, "Commands / configuration used", size=11, bold=True, color=(128, 0, 0))
    for command in assignment["commands"]:
        p = cell.add_paragraph()
        format_paragraph(p, after=1, line=1.0)
        add_text(p, f"• {command}", size=9.5, name="Consolas")
    if assignment["number"] == 1:
        p = cell.add_paragraph()
        format_paragraph(p, before=3, after=0)
        add_text(p, f"Repository: {REPOSITORY}", size=10, bold=True)


def add_picture_fitted(paragraph, image_path, max_width=5.08, max_height=4.9):
    with Image.open(image_path) as image:
        width_px, height_px = image.size
    aspect = width_px / height_px
    width = max_width
    height = width / aspect
    if height > max_height:
        height = max_height
        width = height * aspect
    run = paragraph.add_run()
    run.add_picture(str(image_path), width=Inches(width), height=Inches(height))


def add_step(cell, assignment_number, index, filename, caption, explanation):
    clear_cell(cell)
    folder = SCREENSHOTS / f"assignment-{assignment_number:02d}"
    image_path = folder / filename
    if not image_path.exists():
        raise FileNotFoundError(f"Missing screenshot: {image_path}")

    heading = cell.paragraphs[0]
    format_paragraph(heading, after=1, line=1)
    heading.paragraph_format.keep_with_next = True
    add_text(heading, caption, size=10.5, bold=True)

    description = cell.add_paragraph()
    format_paragraph(description, after=3, line=1.05)
    description.paragraph_format.keep_with_next = True
    add_text(description, explanation, size=10)

    image_paragraph = cell.add_paragraph()
    image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(image_paragraph, after=1, line=1)
    add_picture_fitted(image_paragraph, image_path)

    caption_paragraph = cell.add_paragraph()
    caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_paragraph(caption_paragraph, after=1, line=1)
    add_text(
        caption_paragraph,
        f"Figure {assignment_number}.{index}: {caption}",
        size=8.5,
        italic=True,
    )


def add_result(cell):
    paragraph = clear_cell(cell)
    add_text(
        paragraph,
        "The assignment workflow completed successfully, and the screenshots above were captured from the student's implementation.",
        size=10.5,
    )


def add_assignment(document, assignment, logo_path):
    add_header(document, assignment["number"], logo_path)
    row_count = 10 + len(assignment["steps"]) + 1
    table = document.add_table(rows=row_count, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    for row in table.rows:
        set_cell_width(row.cells[0], 1.88)
        set_cell_width(row.cells[1], 5.22)
        row.cells[0].width = Inches(1.88)
        row.cells[1].width = Inches(5.22)
        for cell in row.cells:
            set_cell_margins(cell)

    labels = [
        "Subject:",
        "TCODE:",
        "Name of Student",
        "PRN No.",
        "Branch",
        "Academic Year &\nSemester",
        "Date of Performance",
        "Title of Assignment:",
        "Theory:",
        "Implementation:",
    ]
    values = [
        "DevOps Lab",
        "TE7950",
        STUDENT_NAME,
        PRN,
        "CSE, Batch (2023-27), L3",
        "2025-26, Sem VII",
        assignment["date"],
        assignment["title"],
    ]

    for index, label in enumerate(labels):
        add_label(table.cell(index, 0), label)
    for index, value in enumerate(values):
        add_value(table.cell(index, 1), value, bold=index in (2, 3, 7))

    add_theory(table.cell(8, 1), assignment)
    add_value(table.cell(9, 1), "Workflow and evidence", bold=True)
    set_cell_shading(table.cell(9, 1), "F7F7F7")

    for index, (filename, caption, explanation) in enumerate(assignment["steps"], start=1):
        row_index = 9 + index
        add_label(table.cell(row_index, 0), f"Step {index}")
        add_step(
            table.cell(row_index, 1),
            assignment["number"],
            index,
            filename,
            caption,
            explanation,
        )

    result_index = row_count - 1
    add_label(table.cell(result_index, 0), "Result:")
    add_result(table.cell(result_index, 1))

    for row in table.rows[:8]:
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.height = Pt(18)

    faculty = document.add_paragraph()
    format_paragraph(faculty, before=5, after=0)
    add_text(faculty, f"Subject Faculty Name: {FACULTY}", size=11)


def configure_document(document):
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.2)
    section.bottom_margin = Inches(0.78)
    section.left_margin = Inches(0.91)
    section.right_margin = Inches(0.80)
    section.header_distance = Inches(0.1)
    section.footer_distance = Inches(0.25)

    normal = document.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(3)

    properties = document.core_properties
    properties.title = "DevOps Lab Record - Shubhankar Sarangi"
    properties.subject = "DevOps Lab - 15 Assignments"
    properties.author = STUDENT_NAME
    properties.last_modified_by = STUDENT_NAME
    properties.keywords = "DevOps, Git, Jira, Jenkins, Docker, Kubernetes, Terraform"
    properties.comments = "Authentic lab workflow record with screenshots."


def main():
    if not LOGO.exists():
        raise FileNotFoundError(f"Institute logo not found: {LOGO}")

    document = Document()
    configure_document(document)

    for index, assignment in enumerate(ASSIGNMENTS):
        if index:
            document.add_page_break()
        add_assignment(document, assignment, LOGO)

    document.save(OUTPUT)
    print(f"Created: {OUTPUT}")
    print(f"Assignments: {len(ASSIGNMENTS)}")
    print(f"Screenshots embedded: {sum(len(a['steps']) for a in ASSIGNMENTS)}")


if __name__ == "__main__":
    main()
