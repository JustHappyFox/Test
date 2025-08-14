#pragma once

#include <string>
#include <unordered_map>
#include <vector>
#include <optional>

struct FireUniform {
    std::string name;
    std::string type; // float, int, vec2, vec3, vec4
    std::vector<float> value; // store as floats; for int cast later
};

struct FirePass {
    std::string name;
    std::string vertexShader;   // GLSL source
    std::string fragmentShader; // GLSL source
};

struct FireEffect {
    std::string name;
    int width = 1280;
    int height = 720;
    std::vector<FireUniform> uniforms;
    std::vector<FirePass> passes; // for now, we will use first pass
};

// Load from a .fire JSON file on disk
std::optional<FireEffect> loadFireFromFile(const std::string& filePath, std::string& outError);