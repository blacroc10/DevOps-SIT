pipeline {
    agent any

    parameters {
        string(
            name: 'GREETING_NAME',
            defaultValue: 'Shubhankar',
            description: 'Name used by the greeting stage'
        )
    }

    stages {
        stage('Greeting') {
            steps {
                echo "Hello ${params.GREETING_NAME}"
            }
        }

        stage('Checkout') {
            steps {
                checkout scm
            }
        }

        stage('Build Docker Image') {
            steps {
                sh 'docker build -t devops-sit-flask:${BUILD_NUMBER} my-flask-app'
            }
        }

        stage('Test') {
            steps {
                sh 'docker image inspect devops-sit-flask:${BUILD_NUMBER} >/dev/null'
            }
        }
    }
}
