-- insert_pdf.lua
-- Pandoc Lua filter to replace PDF insertion placeholders with LaTeX includepdf commands

-- Get the input directory from Pandoc's global variables
local input_dir = ""

function setup_input_dir()
    -- Try to get the input directory from PANDOC_STATE
    if PANDOC_STATE and PANDOC_STATE.input_files and PANDOC_STATE.input_files[1] then
        local input_file = PANDOC_STATE.input_files[1]
        input_dir = input_file:match("(.*/)")
        if not input_dir then
            input_dir = input_file:match("(.*\\)")
        end
        if not input_dir then
            input_dir = ""
        end
    end
end

-- Initialize the input directory
setup_input_dir()

function resolve_path(pdf_path, base_dir)
    -- If it's already an absolute path (starts with drive letter on Windows or / on Unix), use as-is
    if pdf_path:match("^[A-Za-z]:") or pdf_path:match("^/") then
        return pdf_path
    end
    
    -- For relative paths, prepend the base directory
    if base_dir and base_dir ~= "" then
        -- Ensure base_dir ends with a separator
        if not base_dir:match("[/\\]$") then
            base_dir = base_dir .. "/"
        end
        return base_dir .. pdf_path
    else
        return pdf_path
    end
end

function Para(para)
    -- Convert paragraph content to plain text string
    local text = pandoc.utils.stringify(para)
    
    -- Check if the paragraph matches the PDF insertion placeholder pattern
    -- Pattern: [[INSERT: path/to/file.pdf]]
    local pdf_path = text:match("%[%[INSERT:%s*(.-)%s*%]%]")
    
    if pdf_path then
        -- Resolve relative paths based on input document directory
        local resolved_path = resolve_path(pdf_path, input_dir)
        
        -- Convert Windows backslashes to forward slashes for LaTeX compatibility
        local latex_path = resolved_path:gsub("\\", "/")
        
        -- Create LaTeX command to include the PDF
        -- \newpage ensures the PDF starts on a new page
        -- \includepdf[pages=-] includes all pages of the PDF
        local latex_command = "\\newpage\n\\includepdf[pages=-]{" .. latex_path .. "}"
        
        -- Return a RawBlock with LaTeX format
        return pandoc.RawBlock("latex", latex_command)
    end
    
    -- If no match found, return the paragraph unchanged
    return para
end

-- Metadata to describe the filter
local meta = {
    name = "PDF Insertion Filter",
    description = "Replaces [[INSERT: path]] placeholders with LaTeX includepdf commands, resolving relative paths",
    version = "2.0"
}

return {
    {Para = Para},
    meta = meta
}
