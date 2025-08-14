#pragma once

#include <string>
#include <memory>
#include <optional>

#include <GL/glew.h>
#include <GLFW/glfw3.h>

#include "ShaderProgram.h"
#include "Mesh.h"
#include "FireFormat.h"

class FireRuntime {
public:
    FireRuntime();
    ~FireRuntime();

    bool init(const char* title, int width, int height, std::string& outError);
    void shutdown();

    bool loadEffect(const FireEffect& effect, std::string& outError);
    bool loadEffectFromPath(const std::string& filePath, std::string& outError);
    void runMainLoop();

private:
    GLFWwindow* window;
    int windowWidth;
    int windowHeight;

    bool imguiInitialized;

    ShaderProgram shader;
    FullscreenQuadMesh quad;
    FireEffect currentEffect;

    double startTime;
    std::string lastLoadedPath;

    bool initImGui();
    void shutdownImGui();
    void renderFrame();
};