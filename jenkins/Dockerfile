FROM jenkins/jenkins:lts-jdk17

USER root
ARG DOCKER_CLI_VERSION=29.0.1
RUN curl -fsSL \
      "https://download.docker.com/linux/static/stable/x86_64/docker-${DOCKER_CLI_VERSION}.tgz" \
      -o /tmp/docker.tgz \
    && tar -xzf /tmp/docker.tgz -C /tmp \
    && mv /tmp/docker/docker /usr/local/bin/docker \
    && chmod +x /usr/local/bin/docker \
    && rm -rf /tmp/docker /tmp/docker.tgz

USER jenkins
RUN jenkins-plugin-cli --plugins \
    git \
    workflow-aggregator \
    job-dsl
