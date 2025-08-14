#include "Mesh.h"

static const float QUAD_VERTICES[] = {
    // pos.x pos.y pos.z   uv.x uv.y
    -1.0f, -1.0f, 0.0f,    0.0f, 0.0f,
     1.0f, -1.0f, 0.0f,    1.0f, 0.0f,
     1.0f,  1.0f, 0.0f,    1.0f, 1.0f,

    -1.0f, -1.0f, 0.0f,    0.0f, 0.0f,
     1.0f,  1.0f, 0.0f,    1.0f, 1.0f,
    -1.0f,  1.0f, 0.0f,    0.0f, 1.0f,
};

FullscreenQuadMesh::FullscreenQuadMesh()
    : vao(0), vbo(0) {
    glGenVertexArrays(1, &vao);
    glGenBuffers(1, &vbo);

    glBindVertexArray(vao);
    glBindBuffer(GL_ARRAY_BUFFER, vbo);
    glBufferData(GL_ARRAY_BUFFER, sizeof(QUAD_VERTICES), QUAD_VERTICES, GL_STATIC_DRAW);

    // position
    glEnableVertexAttribArray(0);
    glVertexAttribPointer(0, 3, GL_FLOAT, GL_FALSE, 5 * sizeof(float), (void*)0);
    // uv
    glEnableVertexAttribArray(1);
    glVertexAttribPointer(1, 2, GL_FLOAT, GL_FALSE, 5 * sizeof(float), (void*)(3 * sizeof(float)));

    glBindVertexArray(0);
    glBindBuffer(GL_ARRAY_BUFFER, 0);
}

FullscreenQuadMesh::~FullscreenQuadMesh() {
    if (vbo) glDeleteBuffers(1, &vbo);
    if (vao) glDeleteVertexArrays(1, &vao);
}

FullscreenQuadMesh::FullscreenQuadMesh(FullscreenQuadMesh&& other) noexcept {
    vao = other.vao;
    vbo = other.vbo;
    other.vao = 0;
    other.vbo = 0;
}

FullscreenQuadMesh& FullscreenQuadMesh::operator=(FullscreenQuadMesh&& other) noexcept {
    if (this != &other) {
        if (vbo) glDeleteBuffers(1, &vbo);
        if (vao) glDeleteVertexArrays(1, &vao);
        vao = other.vao;
        vbo = other.vbo;
        other.vao = 0;
        other.vbo = 0;
    }
    return *this;
}

void FullscreenQuadMesh::draw() const {
    glBindVertexArray(vao);
    glDrawArrays(GL_TRIANGLES, 0, 6);
    glBindVertexArray(0);
}