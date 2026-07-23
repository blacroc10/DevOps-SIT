terraform {
  required_providers {
    local = {
      source  = "hashicorp/local"
      version = "~> 2.5"
    }
  }
}

provider "local" {}

resource "local_file" "devops_lab" {
  filename = "${path.module}/devops_output.txt"
  content  = "DevOps Infrastructure created using Terraform by Shubhankar Sarangi."
}

output "created_file" {
  value = local_file.devops_lab.filename
}
