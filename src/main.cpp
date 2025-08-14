#include <string>
#include <iostream>

#include "FireFormat.h"
#include "FireRuntime.h"

int main(int argc, char** argv) {
    std::string error;

    FireRuntime runtime;
    if (!runtime.init("Fire Player", 1280, 720, error)) {
        std::cerr << "Init error: " << error << "\n";
        return 1;
    }

    FireEffect effect;
    if (argc >= 2) {
        auto loaded = loadFireFromFile(argv[1], error);
        if (!loaded) {
            std::cerr << "Failed to load .fire: " << error << "\n";
            return 2;
        }
        effect = std::move(*loaded);
    } else {
        // Built-in default effect
        effect.name = "Default";
        effect.width = 1280;
        effect.height = 720;
        effect.passes.push_back({
            "main",
            // Vertex shader
            R"(#version 330 core
layout (location = 0) in vec3 aPos;
layout (location = 1) in vec2 aUV;

out vec2 vUV;

void main() {
    vUV = aUV;
    gl_Position = vec4(aPos, 1.0);
}
)",
            // Fragment shader
            R"(#version 330 core
out vec4 FragColor;

in vec2 vUV;

uniform float iTime;
uniform vec2 iResolution;

float sdCircle(vec2 p, float r) { return length(p) - r; }

void main() {
    vec2 uv = (vUV * 2.0 - 1.0) * vec2(iResolution.x / iResolution.y, 1.0);
    float t = iTime;
    float d = sdCircle(uv + 0.3 * vec2(sin(t * 1.1), cos(t * 1.7)), 0.4 + 0.1 * sin(t * 2.0));
    float glow = exp(-10.0 * abs(d));
    vec3 col = vec3(0.1, 0.2, 0.9) * glow + vec3(0.9, 0.3, 0.1) * smoothstep(0.0, 0.01, -d);
    FragColor = vec4(col, 1.0);
}
)"
        });
    }

    if (!runtime.loadEffect(effect, error)) {
        std::cerr << "Effect error: " << error << "\n";
        return 3;
    }

    runtime.runMainLoop();
    return 0;
}