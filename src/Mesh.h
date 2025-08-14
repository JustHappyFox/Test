#pragma once

#include <GL/glew.h>

class FullscreenQuadMesh {
public:
    FullscreenQuadMesh();
    ~FullscreenQuadMesh();

    FullscreenQuadMesh(const FullscreenQuadMesh&) = delete;
    FullscreenQuadMesh& operator=(const FullscreenQuadMesh&) = delete;

    FullscreenQuadMesh(FullscreenQuadMesh&& other) noexcept;
    FullscreenQuadMesh& operator=(FullscreenQuadMesh&& other) noexcept;

    void draw() const;

private:
    GLuint vao;
    GLuint vbo;
};