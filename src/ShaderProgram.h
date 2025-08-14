#pragma once

#include <string>
#include <unordered_map>
#include <optional>
#include <vector>

#include <GL/glew.h>

class ShaderProgram {
public:
    ShaderProgram();
    ~ShaderProgram();

    ShaderProgram(const ShaderProgram&) = delete;
    ShaderProgram& operator=(const ShaderProgram&) = delete;

    ShaderProgram(ShaderProgram&& other) noexcept;
    ShaderProgram& operator=(ShaderProgram&& other) noexcept;

    bool compileFromSource(const std::string& vertexSource,
                           const std::string& fragmentSource,
                           std::string& outCompileLog);

    void use() const;
    unsigned int id() const { return programId; }

    // Uniform setters (no-ops if uniform not found)
    void setUniformFloat(const char* name, float value) const;
    void setUniformInt(const char* name, int value) const;
    void setUniformVec2(const char* name, float x, float y) const;
    void setUniformVec3(const char* name, float x, float y, float z) const;
    void setUniformVec4(const char* name, float x, float y, float z, float w) const;
    void setUniformMat4(const char* name, const float* mat4ColumnMajor) const;

private:
    unsigned int programId;
    mutable std::unordered_map<std::string, int> uniformCache;

    unsigned int compileShader(unsigned int type, const std::string& source, std::string& logBuffer);
    int getUniformLocation(const char* name) const;
    void destroy();
};