#include "ShaderProgram.h"

#include <sstream>

static bool hasVersionDirective(const std::string& src) {
    return src.find("#version") != std::string::npos;
}

ShaderProgram::ShaderProgram()
    : programId(0) {
}

ShaderProgram::~ShaderProgram() {
    destroy();
}

ShaderProgram::ShaderProgram(ShaderProgram&& other) noexcept {
    programId = other.programId;
    uniformCache = std::move(other.uniformCache);
    other.programId = 0;
}

ShaderProgram& ShaderProgram::operator=(ShaderProgram&& other) noexcept {
    if (this != &other) {
        destroy();
        programId = other.programId;
        uniformCache = std::move(other.uniformCache);
        other.programId = 0;
    }
    return *this;
}

unsigned int ShaderProgram::compileShader(unsigned int type, const std::string& source, std::string& logBuffer) {
    std::string finalSource;
    if (!hasVersionDirective(source)) {
        finalSource = "#version 330 core\n" + source;
    } else {
        finalSource = source;
    }

    unsigned int id = glCreateShader(type);
    const char* srcPtr = finalSource.c_str();
    glShaderSource(id, 1, &srcPtr, nullptr);
    glCompileShader(id);

    int result = 0;
    glGetShaderiv(id, GL_COMPILE_STATUS, &result);
    if (result == GL_FALSE) {
        int length = 0;
        glGetShaderiv(id, GL_INFO_LOG_LENGTH, &length);
        std::string message(length, '\0');
        glGetShaderInfoLog(id, length, &length, message.data());
        logBuffer += message;
        glDeleteShader(id);
        return 0;
    }

    return id;
}

bool ShaderProgram::compileFromSource(const std::string& vertexSource,
                                      const std::string& fragmentSource,
                                      std::string& outCompileLog) {
    outCompileLog.clear();
    destroy();

    unsigned int vs = compileShader(GL_VERTEX_SHADER, vertexSource, outCompileLog);
    unsigned int fs = compileShader(GL_FRAGMENT_SHADER, fragmentSource, outCompileLog);
    if (vs == 0 || fs == 0) {
        if (vs) glDeleteShader(vs);
        if (fs) glDeleteShader(fs);
        return false;
    }

    programId = glCreateProgram();
    glAttachShader(programId, vs);
    glAttachShader(programId, fs);
    glLinkProgram(programId);

    glDeleteShader(vs);
    glDeleteShader(fs);

    int linkStatus = 0;
    glGetProgramiv(programId, GL_LINK_STATUS, &linkStatus);
    if (linkStatus == GL_FALSE) {
        int length = 0;
        glGetProgramiv(programId, GL_INFO_LOG_LENGTH, &length);
        std::string message(length, '\0');
        glGetProgramInfoLog(programId, length, &length, message.data());
        outCompileLog += message;
        glDeleteProgram(programId);
        programId = 0;
        return false;
    }

    uniformCache.clear();
    return true;
}

void ShaderProgram::use() const {
    glUseProgram(programId);
}

int ShaderProgram::getUniformLocation(const char* name) const {
    auto it = uniformCache.find(name);
    if (it != uniformCache.end()) return it->second;
    int location = glGetUniformLocation(programId, name);
    uniformCache[name] = location;
    return location;
}

void ShaderProgram::setUniformFloat(const char* name, float value) const {
    int loc = getUniformLocation(name);
    if (loc >= 0) glUniform1f(loc, value);
}

void ShaderProgram::setUniformInt(const char* name, int value) const {
    int loc = getUniformLocation(name);
    if (loc >= 0) glUniform1i(loc, value);
}

void ShaderProgram::setUniformVec2(const char* name, float x, float y) const {
    int loc = getUniformLocation(name);
    if (loc >= 0) glUniform2f(loc, x, y);
}

void ShaderProgram::setUniformVec3(const char* name, float x, float y, float z) const {
    int loc = getUniformLocation(name);
    if (loc >= 0) glUniform3f(loc, x, y, z);
}

void ShaderProgram::setUniformVec4(const char* name, float x, float y, float z, float w) const {
    int loc = getUniformLocation(name);
    if (loc >= 0) glUniform4f(loc, x, y, z, w);
}

void ShaderProgram::setUniformMat4(const char* name, const float* mat4ColumnMajor) const {
    int loc = getUniformLocation(name);
    if (loc >= 0) glUniformMatrix4fv(loc, 1, GL_FALSE, mat4ColumnMajor);
}

void ShaderProgram::destroy() {
    if (programId != 0) {
        glDeleteProgram(programId);
        programId = 0;
    }
    uniformCache.clear();
}