classdef growth_analysis_module
    % GROWTH_ANALYSIS_MODULE - Contains all bacterial growth analysis functions
    % This module handles data preparation, OD/Luciferase analysis, and enhanced statistics
    
    methods(Static)
        
      function prepareDataGUI(mainFig)
    % prepareDataGUI - Flexible data preparation for 1-4 numeric blocks
    % Automatically detects all numeric blocks and lets user choose which to use

    handles = guidata(mainFig);
    set(handles.statusText,'String','Step 1: Preparing data...');

    % --- select Excel file ---
    [file, path] = uigetfile('*.xlsx','Select Excel File');
    if isequal(file,0)
        set(handles.statusText,'String','No file selected. Ready.');
        return;
    end
    filename = fullfile(path,file);
    set(handles.statusText,'String',['Selected: ' filename newline 'Reading and analyzing...']);

    try
        %% 1) Read entire sheet
        data = readcell(filename);
        [nRows, nCols] = size(data);

        % Basic validation: ensure data is not empty
        if isempty(data)
            error('The selected Excel file is empty or could not be read.');
        end

        %% 2) Detect ALL numeric blocks (not just 2)
        isNum = cellfun(@(x) isnumeric(x) && ~isnan(x), data);
        numCount = sum(isNum, 2);
        dataRows = numCount >= (nCols/2); % Rows with significant numeric content
        
        % Find connected components (consecutive data rows)
        runs = bwconncomp(dataRows);
        lengths = cellfun(@numel, runs.PixelIdxList);
        
        % Sort by size (largest first)
        [sortedLengths, idx] = sort(lengths, 'descend');
        
        % Only keep blocks with reasonable size (at least 3 rows)
        validBlocks = sortedLengths >= 3;
        sortedLengths = sortedLengths(validBlocks);
        idx = idx(validBlocks);
        
        numBlocks = length(sortedLengths);
        
        if numBlocks == 0
            error('No numeric data blocks found in the selected file.');
        end
        
        set(handles.statusText,'String',sprintf('Found %d numeric data blocks. Opening selection window...', numBlocks));
        drawnow;
        
        %% 3) Extract all blocks and their preview data
        allBlocks = cell(1, numBlocks);
        blockPreviews = cell(1, numBlocks);
        blockInfo = cell(1, numBlocks);
        
        for i = 1:numBlocks
            run = runs.PixelIdxList{idx(i)};
            
            % Extract block until blank row
            start_row = min(run);
            end_row = nRows;
            for r = start_row:nRows
                if all(cellfun(@isBlankCell, data(r,:), 'UniformOutput', true))
                    end_row = r-1;
                    break;
                end
            end
            
            blockData = data(start_row:end_row, :);
            allBlocks{i} = blockData;
            
            % Initialize block info structure with all fields explicitly
            numRows = end_row - start_row + 1;
            numCols = size(blockData, 2);
            
            blockInfo{i} = struct(...
                'startRow', start_row, ...
                'endRow', end_row, ...
                'numRows', numRows, ...
                'numCols', numCols, ...
                'medianValue', NaN, ...
                'meanValue', NaN, ...
                'maxValue', NaN, ...
                'minValue', NaN);
            
            % Get numeric preview (skip potential headers)
            previewStart = min(3, size(blockData, 1)); % Skip first 2 rows which might be headers
            if previewStart <= size(blockData, 1) && size(blockData, 2) > 1
                numericData = extractNumericData(blockData(previewStart:end, 2:end));
                if ~isempty(numericData)
                    blockInfo{i}.medianValue = median(numericData(:), 'omitnan');
                    blockInfo{i}.meanValue = mean(numericData(:), 'omitnan');
                    blockInfo{i}.maxValue = max(numericData(:), [], 'omitnan');
                    blockInfo{i}.minValue = min(numericData(:), [], 'omitnan');
                end
            end
            
            % Create preview string with safe access to fields
            medianVal = blockInfo{i}.medianValue;
            minVal = blockInfo{i}.minValue;
            maxVal = blockInfo{i}.maxValue;
            
            if isnan(medianVal)
                blockPreviews{i} = sprintf('Block %d: Rows %d-%d (%d rows)\nNo numeric data found', ...
                    i, start_row, end_row, numRows);
            else
                blockPreviews{i} = sprintf('Block %d: Rows %d-%d (%d rows)\nMedian: %.3f, Range: %.3f-%.3f', ...
                    i, start_row, end_row, numRows, medianVal, minVal, maxVal);
            end
        end
        
        %% 4) Show block selection GUI
        result = showBlockSelectionGUI(blockPreviews, blockInfo, numBlocks);
        
        if isempty(result) || result.cancelled
            set(handles.statusText,'String','Data preparation cancelled. Ready.');
            return;
        end
        
        %% 5) Process selected blocks
        selectedBlocks = result.selectedBlocks;
        blockLabels = result.blockLabels;
        
        % Clean selected blocks
        processedBlocks = cell(1, length(selectedBlocks));
        for i = 1:length(selectedBlocks)
            blockIdx = selectedBlocks(i);
            rawBlock = allBlocks{blockIdx};
            
            % Clean block (remove rows 1 & 3, convert missing)
            cleanBlock = rawBlock;
            if size(cleanBlock, 1) >= 3
                cleanBlock([1,3], :) = [];  % drop rows 1 & 3
            end
            cleanBlock = cellfun(@(x) convertMissing(x), cleanBlock, 'UniformOutput', false);
            processedBlocks{i} = cleanBlock;
        end
        
        %% 6) Save processed blocks and create normalized data
        outputFiles = {};
        outputFolders = {};
        
        % Count how many of each type we have
        odCount = 0;
        lucCount = 0;
        odBlockIdx = [];
        lucBlockIdx = [];
        
        % Generate base timestamp for this session
        baseTimestamp = datestr(now, 'yyyymmdd_HHMMSS');
        
        for i = 1:length(blockLabels)
            label = blockLabels{i};
            
            % Normalize label (LUC -> Luciferase)
            if strcmpi(label, 'LUC')
                label = 'Luciferase';
                blockLabels{i} = label;  % Update the stored label too
            end
            
            % Count and track indices
            if strcmpi(label, 'OD')
                odCount = odCount + 1;
                odBlockIdx = [odBlockIdx, i];
                
                if odCount == 1
                    % First OD block - save in main folder with standard name
                    output_file = fullfile(path, 'ODgrowthAN.xlsx');
                    writecell(processedBlocks{i}, output_file);
                    outputFiles{end+1} = 'ODgrowthAN.xlsx';
                else
                    % Additional OD blocks - create separate folders
                    folder_name = sprintf('OD_Block_%d_%s', odCount, baseTimestamp);
                    folder_path = fullfile(path, folder_name);
                    if ~exist(folder_path, 'dir')
                        mkdir(folder_path);
                    end
                    output_file = fullfile(folder_path, 'ODgrowthAN.xlsx');
                    writecell(processedBlocks{i}, output_file);
                    outputFolders{end+1} = [folder_name '/ODgrowthAN.xlsx'];
                end
                
            elseif strcmpi(label, 'Luciferase')
                lucCount = lucCount + 1;
                lucBlockIdx = [lucBlockIdx, i];
                
                if lucCount == 1
                    % First Luciferase block - save in main folder with standard name
                    output_file = fullfile(path, 'Luciferasenonormalization.xlsx');
                    writecell(processedBlocks{i}, output_file);
                    outputFiles{end+1} = 'Luciferasenonormalization.xlsx';
                else
                    % Additional Luciferase blocks - create separate folders
                    folder_name = sprintf('Luciferase_Block_%d_%s', lucCount, baseTimestamp);
                    folder_path = fullfile(path, folder_name);
                    if ~exist(folder_path, 'dir')
                        mkdir(folder_path);
                    end
                    output_file = fullfile(folder_path, 'Luciferasenonormalization.xlsx');
                    writecell(processedBlocks{i}, output_file);
                    outputFolders{end+1} = [folder_name '/Luciferasenonormalization.xlsx'];
                end
            end
        end
        
        %% 7) Create normalized luciferase if both OD and Luciferase are present
        if ~isempty(odBlockIdx) && ~isempty(lucBlockIdx)
            % Use first OD and first Luciferase blocks for normalization
            odBlock = processedBlocks{odBlockIdx(1)};
            lucBlock = processedBlocks{lucBlockIdx(1)};
            
            % Compute normalized luciferase-to-OD ratio
            newOD = odBlock(2:end, 2:end);
            newLuc = lucBlock(2:end, 2:end);
            ratio = cellfun(@(x,y) safeDiv(x,y), newLuc, newOD, 'UniformOutput', false);
            ratio = cellfun(@(x) max(x,0), ratio, 'UniformOutput', false);
            
            resultMatrix = [
                lucBlock(1,1), odBlock(1,2:end);
                lucBlock(2:end,1), ratio
            ];
            
            writecell(resultMatrix, fullfile(path, 'LuciferaseAN.xlsx'));
        end
        
        %% 8) Final status
        statusMsg = ['Data preparation complete!' newline newline];
        statusMsg = [statusMsg sprintf('Processed %d blocks:', length(selectedBlocks)) newline];
        if odCount > 0
            statusMsg = [statusMsg sprintf('• %d OD block(s)', odCount) newline];
        end
        if lucCount > 0
            statusMsg = [statusMsg sprintf('• %d Luciferase block(s)', lucCount) newline];
        end
        
        if ~isempty(odBlockIdx) && ~isempty(lucBlockIdx)
            statusMsg = [statusMsg newline '✓ Created normalized Luciferase/OD data (using first blocks)' newline];
            outputFiles{end+1} = 'LuciferaseAN.xlsx (normalized)';
        end
        
        statusMsg = [statusMsg newline 'Main output files:' newline];
        for i = 1:length(outputFiles)
            statusMsg = [statusMsg sprintf('• %s', outputFiles{i}) newline];
        end
        
        if ~isempty(outputFolders)
            statusMsg = [statusMsg newline 'Additional blocks in folders:' newline];
            for i = 1:length(outputFolders)
                statusMsg = [statusMsg sprintf('• %s', outputFolders{i}) newline];
            end
            statusMsg = [statusMsg newline 'Note: To analyze additional blocks, copy the desired' newline];
            statusMsg = [statusMsg 'folder contents to main directory and run Steps 2-5.' newline];
        end
        
        set(handles.statusText, 'String', statusMsg);

    catch ME
        set(handles.statusText,'String',['Error: ' ME.message]);
        disp(getReport(ME,'extended'));
    end

    %% Helper Functions (nested inside prepareDataGUI)
    function tf = isBlankCell(x)
        if isnumeric(x)
            tf = isempty(x) || isnan(x);
        elseif isstring(x)
            tf = strlength(x) == 0;
        elseif ischar(x)
            tf = isempty(x);
        elseif ismissing(x)
            tf = true;
        else
            tf = false;
        end
    end

    function numData = extractNumericData(cellData)
        numData = [];
        for i = 1:numel(cellData)
            val = cellData{i};
            if isnumeric(val) && ~isnan(val)
                numData = [numData; val];
            elseif (ischar(val) || isstring(val))
                numVal = str2double(val);
                if ~isnan(numVal)
                    numData = [numData; numVal];
                end
            end
        end
    end

    function output = convertMissing(input)
        if ismissing(input)
            output = '';
        else
            output = input;
        end
    end

    function result = safeDiv(a, b)
        if isempty(a) || isempty(b)
            result = 0;
        else
            a = double(a);
            b = double(b);
            if b == 0
                result = 0;
            else
                result = a / b;
                if result < 0
                    result = 0;
                end
            end
        end
    end

    function result = showBlockSelectionGUI(blockPreviews, blockInfo, numBlocks)
        % Create GUI for block selection and labeling
        
        result = struct();
        result.cancelled = false;
        result.selectedBlocks = [];
        result.blockLabels = {};
        
        % Calculate figure dimensions based on content
        figWidth = 700;
        headerHeight = 140;
        blockHeight = 70;  % Increased for better spacing
        buttonHeight = 60;
        padding = 20;
        
        contentHeight = numBlocks * blockHeight;
        figHeight = headerHeight + contentHeight + buttonHeight + padding * 3;
        figHeight = max(figHeight, 500); % Minimum height
        
        % Create figure
        blockFig = figure('Name', 'Select Data Blocks', ...
                          'Position', [300, 200, figWidth, figHeight], ...
                          'NumberTitle', 'off', 'MenuBar', 'none', ...
                          'WindowStyle', 'modal', 'Resize', 'off');
        
        % Calculate positions from top down
        currentY = figHeight - padding;
        
        % Title
        titleHeight = 35;
        currentY = currentY - titleHeight;
        uicontrol(blockFig, 'Style', 'text', ...
                  'String', 'Select and Label Data Blocks', ...
                  'Position', [padding, currentY, figWidth-2*padding, titleHeight], ...
                  'FontSize', 16, 'FontWeight', 'bold', ...
                  'HorizontalAlignment', 'center');
        
        % Instructions
        instrHeight = 80;
        currentY = currentY - instrHeight - 10;
        instrText = ['Select which data blocks to process and label as "OD" or "Luciferase".' newline newline ...
                    '• First blocks: saved as ODgrowthAN.xlsx & Luciferasenonormalization.xlsx' newline ...
                    '• Additional blocks: saved in separate timestamped folders'];
        uicontrol(blockFig, 'Style', 'text', ...
                  'String', instrText, ...
                  'Position', [padding, currentY, figWidth-2*padding, instrHeight], ...
                  'FontSize', 10, 'HorizontalAlignment', 'left', ...
                  'BackgroundColor', [0.95, 0.95, 0.95]);
        
        % Headers
        headerHeight = 25;
        currentY = currentY - headerHeight - 15;
        
        uicontrol(blockFig, 'Style', 'text', 'String', 'Select', ...
                  'Position', [padding + 10, currentY, 60, headerHeight], ...
                  'FontWeight', 'bold', 'FontSize', 11);
        uicontrol(blockFig, 'Style', 'text', 'String', 'Block Information', ...
                  'Position', [padding + 80, currentY, 250, headerHeight], ...
                  'FontWeight', 'bold', 'FontSize', 11);
        uicontrol(blockFig, 'Style', 'text', 'String', 'Label', ...
                  'Position', [padding + 340, currentY, 80, headerHeight], ...
                  'FontWeight', 'bold', 'FontSize', 11);
        uicontrol(blockFig, 'Style', 'text', 'String', 'Suggestion', ...
                  'Position', [padding + 430, currentY, 200, headerHeight], ...
                  'FontWeight', 'bold', 'FontSize', 11);
        
        % Block selection controls
        checkboxes = cell(1, numBlocks);
        labelEdits = cell(1, numBlocks);
        
        for i = 1:numBlocks
            currentY = currentY - blockHeight;
            
            % Add separator line
            if i > 1
                uicontrol(blockFig, 'Style', 'frame', ...
                          'Position', [padding, currentY + blockHeight - 5, figWidth-2*padding, 1], ...
                          'BackgroundColor', [0.8, 0.8, 0.8]);
            end
            
            % Checkbox
            checkboxes{i} = uicontrol(blockFig, 'Style', 'checkbox', ...
                                      'Position', [padding + 25, currentY + 25, 20, 20], ...
                                      'Value', i <= 2); % Default: select first 2 blocks
            
            % Block info with better formatting
            infoText = blockPreviews{i};
            uicontrol(blockFig, 'Style', 'text', ...
                      'String', infoText, ...
                      'Position', [padding + 80, currentY + 5, 250, blockHeight - 10], ...
                      'HorizontalAlignment', 'left', 'FontSize', 9, ...
                      'BackgroundColor', [0.98, 0.98, 0.98]);
            
            % Label edit box with smart defaults
            defaultLabel = getDefaultLabel(i, blockInfo{i});
            labelEdits{i} = uicontrol(blockFig, 'Style', 'edit', ...
                                      'String', defaultLabel, ...
                                      'Position', [padding + 340, currentY + 22, 80, 25], ...
                                      'FontSize', 10, 'HorizontalAlignment', 'center');
            
            % Suggested labels with better formatting
            suggestions = getSuggestedLabels(blockInfo{i});
            if ~isempty(suggestions)
                uicontrol(blockFig, 'Style', 'text', ...
                          'String', suggestions, ...
                          'Position', [padding + 430, currentY + 20, 200, 30], ...
                          'FontSize', 9, 'ForegroundColor', [0.4, 0.4, 0.8], ...
                          'HorizontalAlignment', 'left');
            end
        end
        
        % Buttons at bottom with proper spacing
        buttonWidth = 120;
        buttonHeight = 35;
        buttonY = padding + 10;
        
        % Center the buttons
        totalButtonWidth = buttonWidth * 2 + 20; % 2 buttons + spacing
        startX = (figWidth - totalButtonWidth) / 2;
        
        uicontrol(blockFig, 'Style', 'pushbutton', 'String', 'Process Selected', ...
                  'Position', [startX, buttonY, buttonWidth, buttonHeight], ...
                  'FontWeight', 'bold', 'FontSize', 11, ...
                  'BackgroundColor', [0.2, 0.7, 0.2], 'ForegroundColor', 'white', ...
                  'Callback', @processCallback);
        
        uicontrol(blockFig, 'Style', 'pushbutton', 'String', 'Cancel', ...
                  'Position', [startX + buttonWidth + 20, buttonY, buttonWidth, buttonHeight], ...
                  'FontSize', 11, 'Callback', @cancelCallback);
        
        % Add status text area
        statusText = uicontrol(blockFig, 'Style', 'text', ...
                              'String', sprintf('Found %d data blocks. Select and label the ones to process.', numBlocks), ...
                              'Position', [padding, buttonY + buttonHeight + 10, figWidth-2*padding, 20], ...
                              'FontSize', 10, 'ForegroundColor', [0.3, 0.3, 0.3], ...
                              'HorizontalAlignment', 'center');
        
        % Wait for user action
        uiwait(blockFig);
        
        function processCallback(~, ~)
            % Get selected blocks and labels
            selectedIdx = [];
            labels = {};
            
            for j = 1:numBlocks
                if get(checkboxes{j}, 'Value')
                    selectedIdx = [selectedIdx, j];
                    label = strtrim(get(labelEdits{j}, 'String'));
                    if isempty(label)
                        label = sprintf('Block%d', j);
                    end
                    labels{end+1} = label;
                end
            end
            
            if isempty(selectedIdx)
                set(statusText, 'String', 'Please select at least one block to process.', ...
                    'ForegroundColor', [0.8, 0.2, 0.2]);
                return;
            end
            
            % Validate labels - only OD and Luciferase allowed
            validLabels = {'OD', 'Luciferase', 'LUC'};  % Accept LUC as synonym
            for j = 1:length(labels)
                if ~any(strcmpi(labels{j}, validLabels))
                    set(statusText, 'String', sprintf('Invalid label "%s". Only "OD" or "Luciferase" allowed.', labels{j}), ...
                        'ForegroundColor', [0.8, 0.2, 0.2]);
                    return;
                end
            end
            
            result.selectedBlocks = selectedIdx;
            result.blockLabels = labels;
            result.cancelled = false;
            
            % Show success message briefly
            set(statusText, 'String', sprintf('Processing %d selected blocks...', length(selectedIdx)), ...
                'ForegroundColor', [0.2, 0.6, 0.2]);
            pause(0.5);
            
            delete(blockFig);
        end

        function cancelCallback(~, ~)
            result.cancelled = true;
            delete(blockFig);
        end
    end

    function defaultLabel = getDefaultLabel(blockIndex, blockInfo)
        % Smart default labeling - only OD or Luciferase
        
        if isnan(blockInfo.medianValue)
            % If no numeric data, guess based on order
            if blockIndex == 1
                defaultLabel = 'OD';
            else
                defaultLabel = 'Luciferase';
            end
            return;
        end
        
        % Use median value to distinguish OD vs Luciferase
        median_val = blockInfo.medianValue;
        
        if median_val < 1
            defaultLabel = 'OD';  % Small values typically OD (0.01 - 3.0 range)
        else
            defaultLabel = 'Luciferase';  % Larger values typically luciferase
        end
    end

    function suggestions = getSuggestedLabels(blockInfo)
        % Suggest OD or Luciferase based on data range
        
        if isnan(blockInfo.medianValue)
            suggestions = 'OD or Luciferase';
            return;
        end
        
        median_val = blockInfo.medianValue;
        max_val = blockInfo.maxValue;
        
        if median_val < 1 && max_val < 5
            suggestions = 'OD (values < 1)';
        elseif median_val > 50
            suggestions = 'Luciferase (large values)';
        elseif median_val > 1 && median_val < 50
            suggestions = 'Likely Luciferase';
        else
            suggestions = 'OD or Luciferase';
        end
    end

end

        
        
        function analyzeODGUI(mainFig)
    growth_analysis_module.analyzeDataGUI(mainFig, 'OD');
        end
        
        function analyzeLuciferaseGUI(mainFig)
    growth_analysis_module.analyzeDataGUI(mainFig, 'LUC');
        end
        
        function analyzeDataGUI(mainFig, dataType)
             handles = guidata(mainFig);
    
    % Update status
    if strcmp(dataType, 'OD')
        set(handles.statusText, 'String', 'Step 2: Analyzing OD growth data...');
        defaultFile = 'ODgrowthAN.xlsx';
        outputFolder = 'odgraphs';
        outputExcel = 'GrahpPad_OD.xlsx';
        yLabel = 'OD 600nm';
        combinedTitle = 'Averages of OD growth';
    else
        set(handles.statusText, 'String', 'Step 3: Analyzing Luciferase data...');
        defaultFile = 'LuciferaseAN.xlsx';
        outputFolder = 'LUCgraphs';
        outputExcel = 'GraphPad_LUC.xlsx';
        yLabel = 'Luciferase';
        combinedTitle = 'Averages of Luciferase';
    end
    
    % Select file
    [filename, filepath] = uigetfile({'*.xlsx;*.xls', 'Excel Files (*.xlsx, *.xls)'}, ...
                            'Select Excel File', defaultFile);
    if isequal(filename, 0)
        set(handles.statusText, 'String', 'No file selected. Ready for next operation.');
        return;
    end
    
    try
        % Read data
        excelFilePath = fullfile(filepath, filename);
        [dataMatrix, textData] = xlsread(excelFilePath);
        
        % Extract time vector
        timeVectorInSeconds = dataMatrix(1, :);
        timeVectorInHours = timeVectorInSeconds / 3600;
        
        % Get number of groups and repetitions per group
        prompt = {'Enter the number of groups (e.g., 2):', 'Enter the number of repetitions per group (e.g., 4):'};
        dlgtitle = 'Group Configuration';
        dims = [1 50];
        definput = {'2', '4'};
        userInput = inputdlg(prompt, dlgtitle, dims, definput);

        if isempty(userInput)
            set(handles.statusText, 'String', 'Operation canceled. Ready for next operation.');
            return;
        end

        numGroups = str2double(userInput{1});
        repetitionsPerGroup = str2double(userInput{2});

        % Input validation for numGroups and repetitionsPerGroup
        if isnan(numGroups) || numGroups < 1 || floor(numGroups) ~= numGroups
            errordlg('Number of groups must be a positive integer.', 'Invalid Input');
            set(handles.statusText, 'String', 'Invalid input. Operation canceled.');
            return;
        end
        if isnan(repetitionsPerGroup) || repetitionsPerGroup < 1 || floor(repetitionsPerGroup) ~= repetitionsPerGroup
            errordlg('Number of repetitions per group must be a positive integer.', 'Invalid Input');
            set(handles.statusText, 'String', 'Invalid input. Operation canceled.');
            return;
        end
        
        % Generate the same unique colors as used in selectWellsGUI
        groupColors = cell(1, numGroups);
        for i = 1:numGroups
            groupColors{i} = growth_analysis_module.generateUniqueColor(i);
        end
        
        % Create the persistent well selection window
        set(handles.statusText, 'String', 'Opening well selection window...');
        drawnow;
        
        % Call the enhanced well selection GUI (it will stay open)
        growth_analysis_module.selectWellsEnhancedGUI(textData(2:end, 1), numGroups, repetitionsPerGroup);
        
        
        
        % Update main status
        set(handles.statusText, 'String', [
            'Well selection window opened!' newline newline ...
            'Instructions:' newline ...
            '1. Select wells for each group in the selection window' newline ...
            '2. Check/uncheck groups to include in plot' newline ...
            '3. Click "Plot Results" in the selection window' newline ...
            '4. Window stays open for multiple plots' newline newline ...
            'Ready for analysis!']);
        
        % Store analysis data for callback access
        analysisData = struct();
        analysisData.dataMatrix = dataMatrix;
        analysisData.textData = textData;
        analysisData.timeVectorInHours = timeVectorInHours;
        analysisData.filepath = filepath;
        analysisData.outputFolder = outputFolder;
        analysisData.outputExcel = outputExcel;
        analysisData.yLabel = yLabel;
        analysisData.combinedTitle = combinedTitle;
        analysisData.dataType = dataType;
        analysisData.groupColors = groupColors;
        analysisData.handles = handles;
        
        % Store analysis data globally
        setappdata(0, 'currentAnalysisData', analysisData);
        
        % Set up a timer to check for plot requests from the well selection window
        if isempty(timerfind('Name', 'WellSelectionMonitor'))
            t = timer('Name', 'WellSelectionMonitor', ...
                     'Period', 0.5, ...
                     'ExecutionMode', 'fixedRate', ...
                     'TimerFcn', @growth_analysis_module.checkForPlotRequest);
            start(t);
        end
        
    catch ME
        set(handles.statusText, 'String', ['Error during analysis setup: ' ME.message newline newline ...
                                          'Please check your input file and try again.']);
        disp(getReport(ME, 'extended'));
    end
    
    % Callback function for manual analysis trigger
    function runAnalysisCallback(~, ~)
        % Get the latest result from the persistent window
        result = growth_analysis_module.getPersistentWellSelectionResult();
        
        if isempty(result)
            msgbox('No well selection found. Please make selections in the well selection window first.', 'No Selection');
            return;
        end
        
        % Run the analysis
        growth_analysis_module.performAnalysisWithResult(result);
    end
end
        
        
        function checkForPlotRequest(~, ~)
 % Check if there's a plot request from the well selection window
    
    figHandle = getappdata(0, 'persistentWellSelectionFig');
    
    if isempty(figHandle) || ~ishandle(figHandle)
        % Well selection window is closed, stop timer
        t = timerfind('Name', 'WellSelectionMonitor');
        if ~isempty(t)
            stop(t);
            delete(t);
        end
        
        % Clean up analysis data
        if isappdata(0, 'currentAnalysisData')
            rmappdata(0, 'currentAnalysisData');
        end
        
        return;
    end
    
    % Check if there's a new result
    result = getappdata(figHandle, 'latestResult');
    lastProcessedTime = getappdata(0, 'lastProcessedTime');
    
    if isempty(lastProcessedTime)
        lastProcessedTime = 0;
        setappdata(0, 'lastProcessedTime', lastProcessedTime);
    end
    
    if ~isempty(result)
        % Get current time from well selection window data
        wellData = getappdata(figHandle, 'wellData');
        if isfield(wellData, 'statusText') && ishandle(wellData.statusText)
            statusString = get(wellData.statusText, 'String');
            
            % Simple check: if status contains "Plot generated", process it
            if contains(statusString, 'Plot generated') && ~contains(statusString, 'processed')
                growth_analysis_module.performAnalysisWithResult(result);
                
                % Mark as processed
                set(wellData.statusText, 'String', [statusString ' (processed)']);
            end
        end
    end       
        end
        
        function performAnalysisWithResult(result)
% Perform the actual analysis with the given result
    
    % Get analysis data
    analysisData = getappdata(0, 'currentAnalysisData');
    if isempty(analysisData)
        msgbox('Analysis data not found. Please restart the analysis.', 'Error');
        return;
    end
    
    try
        % Extract data
        dataMatrix = analysisData.dataMatrix;
        textData = analysisData.textData;
        timeVectorInHours = analysisData.timeVectorInHours;
        filepath = analysisData.filepath;
        outputFolder = analysisData.outputFolder;
        outputExcel = analysisData.outputExcel;
        yLabel = analysisData.yLabel;
        combinedTitle = analysisData.combinedTitle;
        dataType = analysisData.dataType;
        groupColors = analysisData.groupColors;
        handles = analysisData.handles;
        
        % Extract group information from result
        groupNames = result.groupNames;
        groupRows = result.selectedRows;
        numGroups = length(groupNames);
        
        % Initialize arrays
        groupData = cell(1, numGroups);
        groupAverage = cell(1, numGroups);
        groupStd = cell(1, numGroups);
        
        % Calculate data for each group
        for i = 1:numGroups
            if ~isempty(groupRows{i})
                groupData{i} = dataMatrix(groupRows{i} + 1, :);
                groupAverage{i} = mean(groupData{i}, 1);
                groupStd{i} = std(groupData{i}, 0, 1);
            else
                % Handle empty groups
                groupData{i} = [];
                groupAverage{i} = [];
                groupStd{i} = [];
            end
        end
        
        % Create output folder
        outputFolderPath = fullfile(filepath, outputFolder);
        if ~exist(outputFolderPath, 'dir')
            mkdir(outputFolderPath);
        end
        
        % Generate timestamp for unique filenames
        timestamp = datestr(now, 'yyyymmdd_HHMMSS');
        
        % Save data for GraphPad (only for groups with data)
        outputExcelFilePath = fullfile(filepath, sprintf('%s_%s.xlsx', outputExcel(1:end-5), timestamp));
        for i = 1:numGroups
            if ~isempty(groupData{i})
                groupSheetData = [textData(groupRows{i} + 1, 1), num2cell(groupData{i})];
                groupSheetData = [['Time (hours)', num2cell(timeVectorInHours)]; groupSheetData];
                headers = ['Row Name', cell(1, size(groupData{i}, 2))];
                groupSheetData = [headers; groupSheetData];
                writecell(groupSheetData, outputExcelFilePath, 'Sheet', groupNames{i}, 'WriteMode', 'overwrite');
            end
        end
        
        % Color map for groups - use colors based on original group order
        colorMap = zeros(numGroups, 3);
        for i = 1:numGroups
            if i <= length(groupColors)
                colorMap(i, :) = groupColors{i};
            else
                colorMap(i, :) = growth_analysis_module.generateUniqueColor(i);
            end
        end
        
        % Plot individual groups (only for groups with data)
        for i = 1:numGroups
            if ~isempty(groupData{i})
                fig = figure('Name', sprintf('%s_%s', groupNames{i}, timestamp));
                if strcmp(dataType, 'OD')
                    growth_analysis_module.plotODGroup(groupRows{i}, groupData{i}, colorMap(i, :), groupNames{i}, timeVectorInHours, textData);
                else
                    growth_analysis_module.plotLuciferaseGroup(groupRows{i}, groupData{i}, colorMap(i, :), groupNames{i}, timeVectorInHours, textData);
                end
                
                % Save individual plot
                figFilename = sprintf('%s_%s.png', groupNames{i}, timestamp);
                saveas(fig, fullfile(outputFolderPath, figFilename));
            end
        end
        
        % Combined plot (only include groups with data)
        groupsWithData = ~cellfun(@isempty, groupData);
        if any(groupsWithData)
            figCombined = figure('Name', sprintf('%s_%s', combinedTitle, timestamp));
            
            if strcmp(dataType, 'OD')
                % Plot OD with filled regions for std - FIXED: robust fill function
                hold on;
                h = gobjects(1, sum(groupsWithData));
                plotIndex = 1;
                
                % Plot the standard deviations as filled regions
                for i = 1:numGroups
                    if groupsWithData(i)
                        try
                            % FIXED: Ensure proper dimensions and no negative values
                            avgData = groupAverage{i};
                            stdData = groupStd{i};
                            timeData = timeVectorInHours;
                            
                            % Ensure all are row vectors
                            if size(avgData, 1) > size(avgData, 2)
                                avgData = avgData';
                            end
                            if size(stdData, 1) > size(stdData, 2)
                                stdData = stdData';
                            end
                            if size(timeData, 1) > size(timeData, 2)
                                timeData = timeData';
                            end
                            
                            % Remove any NaN or infinite values
                            validIdx = isfinite(avgData) & isfinite(stdData) & isfinite(timeData);
                            if any(validIdx)
                                avgData = avgData(validIdx);
                                stdData = stdData(validIdx);
                                timeData = timeData(validIdx);
                                
                                % Calculate bounds
                                lowerBound = max(0, avgData - stdData);
                                upperBound = avgData + stdData;
                                
                                % Create fill vectors
                                fillX = [timeData, fliplr(timeData)];
                                fillY = [lowerBound, fliplr(upperBound)];
                                
                                % Only fill if we have valid data
                                if length(fillX) > 2 && length(fillY) > 2
                                    % Try fill first, fallback to patch if it fails
                                    try
                                        fill(fillX, fillY, colorMap(i, :), 'FaceAlpha', 0.3, 'EdgeColor', 'none');
                                    catch
                                        % Fallback to patch
                                        patch(fillX, fillY, colorMap(i, :), 'FaceAlpha', 0.3, 'EdgeColor', 'none');
                                    end
                                end
                            end
                        catch ME
                            fprintf('Fill error for group %d: %s\n', i, ME.message);
                            % Continue without fill for this group
                        end
                    end
                end
                
                % Plot the averages
                for i = 1:numGroups
                    if groupsWithData(i)
                        h(plotIndex) = plot(timeVectorInHours, groupAverage{i}, 'LineWidth', 1.5, 'MarkerSize', 8, 'Color', colorMap(i, :));
                        plotIndex = plotIndex + 1;
                    end
                end
            else
                % Plot Luciferase with error bands - FIXED: robust fill function
                hold on;
                h = gobjects(1, sum(groupsWithData));
                plotIndex = 1;
                
                % Plot averages with error bands
                for i = 1:numGroups
                    if groupsWithData(i)
                        try
                            % FIXED: Ensure proper dimensions and no negative values
                            avgData = groupAverage{i};
                            stdData = groupStd{i};
                            timeData = timeVectorInHours;
                            
                            % Ensure all are row vectors
                            if size(avgData, 1) > size(avgData, 2)
                                avgData = avgData';
                            end
                            if size(stdData, 1) > size(stdData, 2)
                                stdData = stdData';
                            end
                            if size(timeData, 1) > size(timeData, 2)
                                timeData = timeData';
                            end
                            
                            % Remove any NaN or infinite values
                            validIdx = isfinite(avgData) & isfinite(stdData) & isfinite(timeData);
                            if any(validIdx)
                                avgData = avgData(validIdx);
                                stdData = stdData(validIdx);
                                timeData = timeData(validIdx);
                                
                                % Calculate bounds
                                lowerBound = max(0, avgData - stdData);
                                upperBound = avgData + stdData;
                                
                                % Create fill vectors
                                fillX = [timeData, fliplr(timeData)];
                                fillY = [lowerBound, fliplr(upperBound)];
                                
                                % Only fill if we have valid data
                                if length(fillX) > 2 && length(fillY) > 2
                                    % Try fill first, fallback to patch if it fails
                                    try
                                        fill(fillX, fillY, colorMap(i, :), 'FaceAlpha', 0.3, 'EdgeColor', 'none');
                                    catch
                                        % Fallback to patch
                                        patch(fillX, fillY, colorMap(i, :), 'FaceAlpha', 0.3, 'EdgeColor', 'none');
                                    end
                                end
                            end
                        catch ME
                            fprintf('Fill error for group %d: %s\n', i, ME.message);
                            % Continue without fill for this group
                        end
                    end
                end
                % Plot the averages
                for i = 1:numGroups
                    if groupsWithData(i)
                        h(plotIndex) = plot(timeVectorInHours, groupAverage{i}, 'LineWidth', 1.5, 'MarkerSize', 8, 'Color', colorMap(i, :));
                        plotIndex = plotIndex + 1;
                    end
                end
            end
            
            % Formatting for combined plot
            title(combinedTitle, 'FontSize', 14, 'FontWeight', 'bold');
            xlabel('Time (hours)', 'FontSize', 12, 'FontWeight', 'bold');
            ylabel(yLabel, 'FontSize', 12, 'FontWeight', 'bold');
            legend(h, groupNames(groupsWithData), 'Location', 'best', 'FontSize', 10);
            grid on;
            set(gca, 'FontSize', 10, 'LineWidth', 1.5);
            hold off;
            
            % Save combined plot
            combinedFilename = sprintf('%s_%s.png', combinedTitle, timestamp);
            saveas(figCombined, fullfile(outputFolderPath, combinedFilename));
        end

        %% ========================================================================
%% STEP 1: MODIFY performAnalysisWithResult FUNCTION
%% ========================================================================

% FIND your existing performAnalysisWithResult function and REPLACE the last part
% (after the combined plot section) with this:

       
        
        % MODIFIED: Store data for enhanced analysis (don't launch automatically)
        enhancedAnalysisData = struct();
        enhancedAnalysisData.groupData = groupData;
        enhancedAnalysisData.groupAverage = groupAverage;
        enhancedAnalysisData.groupStd = groupStd;
        enhancedAnalysisData.groupNames = groupNames;
        enhancedAnalysisData.timeVectorInHours = timeVectorInHours;
        enhancedAnalysisData.dataType = dataType;
        enhancedAnalysisData.outputFolderPath = outputFolderPath;
        enhancedAnalysisData.timestamp = timestamp;
        enhancedAnalysisData.textData = textData;
        enhancedAnalysisData.groupRows = groupRows;
        
        % Store globally for button access
        setappdata(0, 'enhancedAnalysisData', enhancedAnalysisData);
        
        % Update main GUI status
        set(handles.statusText, 'String', [
            dataType ' analysis completed successfully!' newline newline ...
            'Results saved to: ' outputFolderPath newline newline ...
            'Generated files:' newline ...
            '• Individual group plots (' strjoin(groupNames, ', ') ')' newline ...
            '• Combined average plot' newline ...
            '• GraphPad-compatible Excel file' newline newline ...
            'Click "Enhanced Analysis" button for detailed statistical analysis!' newline ...
            'Well selection window remains open for additional analyses.']);
        
    catch ME
        set(handles.statusText, 'String', ['Error during analysis: ' ME.message newline newline ...
                                          'Please check your selections and try again.']);
        disp(getReport(ME, 'extended'));
    end       
        end
        
        function plotODGroup(groupRows, groupData, groupColor, groupName, timeVector, textData)
hold on;
    for i = 1:length(groupRows)
        h = plot(timeVector, groupData(i, :), 'LineWidth', 2, 'Color', groupColor);
        set(h, 'DisplayName', sprintf('%s - %s', groupName, textData{groupRows(i) + 1, 1}));
    end
    hold off;
    
    xlabel('Time (hours)', 'FontSize', 12, 'FontWeight', 'bold');
    ylabel('OD 600nm', 'FontSize', 12, 'FontWeight', 'bold');
    title(groupName, 'FontSize', 14, 'FontWeight', 'bold');
    grid on;
    set(gca, 'FontSize', 10, 'LineWidth', 1.5);
    
    % Enable data cursor mode
    dcm = datacursormode(gcf);
    set(dcm, 'UpdateFcn', @growth_analysis_module.customDataCursor);
        end
        
        function plotLuciferaseGroup(groupRows, groupData, groupColor, groupName, timeVector, textData)
hold on;
    for i = 1:length(groupRows)
        h = plot(timeVector, groupData(i, :), 'LineWidth', 2, 'Color', groupColor);
        set(h, 'DisplayName', sprintf('%s - %s', groupName, textData{groupRows(i) + 1, 1}));
    end
    hold off;
    
    xlabel('Time (hours)', 'FontSize', 12, 'FontWeight', 'bold');
    ylabel('Luciferase', 'FontSize', 12, 'FontWeight', 'bold');
    title(groupName, 'FontSize', 14, 'FontWeight', 'bold');
    grid on;
    set(gca, 'FontSize', 10, 'LineWidth', 1.5);
    
    % Enable data cursor mode
    dcm = datacursormode(gcf);
    set(dcm, 'UpdateFcn', @growth_analysis_module.customDataCursor);
        end
        
        function output_txt = customDataCursor(~, event_obj)
pos = get(event_obj, 'Position');
    output_txt = {sprintf('Time: %.2f hours', pos(1)), sprintf('Y: %.2f', pos(2))};
    
    % Retrieve the display name set for the line
    display_name = get(event_obj.Target, 'DisplayName');
    output_txt{end+1} = ['Name: ' display_name];
        end
        
        function color = generateUniqueColor(index)
% Generate unique colors for any number of groups
    % Uses HSV color space to create visually distinct colors
    
    if index <= 10
        % Use predefined colors for first 10 groups
        predefinedColors = {
            [0.8, 0.2, 0.2], % Red
            [0.2, 0.6, 0.8], % Blue
            [0.2, 0.8, 0.2], % Green
            [0.8, 0.8, 0.2], % Yellow
            [0.8, 0.2, 0.8], % Magenta
            [0.2, 0.8, 0.8], % Cyan
            [1.0, 0.5, 0.0], % Orange
            [0.5, 0.0, 0.5], % Purple
            [0.5, 0.5, 0.0], % Olive
            [0.0, 0.5, 0.5]  % Teal
        };
        color = predefinedColors{index};
    else
        % Generate unique colors using HSV color space
        % Distribute hues evenly around the color wheel
        hue = mod((index - 11) * 0.618033988749895, 1); % Golden ratio for even distribution
        
        % Vary saturation and value to create more distinct colors
        saturation = 0.7 + 0.3 * sin(index * 1.3); % Between 0.7 and 1.0
        value = 0.6 + 0.4 * cos(index * 1.7);      % Between 0.6 and 1.0
        
        % Convert HSV to RGB
        color = hsv2rgb([hue, saturation, value]);
    end
    
    % Ensure color values are in valid range [0, 1]
    color = max(0, min(1, color));
        end
        
        function result = getPersistentWellSelectionResult()
% Get the latest result from the persistent well selection window
    figHandle = getappdata(0, 'persistentWellSelectionFig');
    
    if isempty(figHandle) || ~ishandle(figHandle)
        result = [];
        return;
    end
    
    result = getappdata(figHandle, 'latestResult');
        end
        
        function selectedRows = selectWellsEnhancedGUI(rowLabels, numGroups, repetitionsPerGroup)
% Enhanced well selection GUI with group checkboxes for plot inclusion
    % 
    % Inputs:
    %   rowLabels - Cell array of labels for rows
    %   numGroups - Number of groups to assign wells to
    %   repetitionsPerGroup - Expected number of repetitions per group
    %
    % Output:
    %   selectedRows - Cell array with numGroups elements, each containing indices of
    %                  selected wells for the corresponding group
    
    % Input validation
    if nargin < 2
        numGroups = 1;
    end
    if nargin < 3
        repetitionsPerGroup = 3;
    end
    
    % Generate unique colors for each group
    groupColors = cell(1, numGroups);
    for i = 1:numGroups
        groupColors{i} = growth_analysis_module.generateUniqueColor(i);
    end
    
    % Create figure with enhanced layout
    wellFig = figure('Name', 'Enhanced Well Selection', 'Position', [200, 200, 1200, 900], ...
                    'NumberTitle', 'off', 'MenuBar', 'none');
    
    % Create panel for plate layout
    platePanel = uipanel(wellFig, 'Title', 'Plate Layout', 'Position', [0.05, 0.25, 0.60, 0.65]);
    
    % Create panel for group information with scrollable content
    groupPanel = uipanel(wellFig, 'Title', 'Group Information', 'Position', [0.68, 0.25, 0.28, 0.65]);
    
    % Create scrollable panel inside group panel
    % Initial size for scrollPanel, will be adjusted based on content
    scrollPanel = uipanel(groupPanel, 'Position', [0.02, 0.02, 0.96, 0.96], 'BorderType', 'none');
    
    % Create info text with repetition information
    infoText = uicontrol(wellFig, 'Style', 'text', ...
        'String', sprintf('Select %d wells for Group 1 of %d (Expected: %d repetitions)', ...
                         repetitionsPerGroup, numGroups, repetitionsPerGroup), ...
        'Position', [160, 750, 500, 30], 'FontSize', 12, 'FontWeight', 'bold');
    
    % Dedicated status text for well selection GUI
    wellStatusText = uicontrol(wellFig, 'Style', 'text', ...
        'String', 'Ready.', ...
        'Position', [50, 200, 700, 30], 'HorizontalAlignment', 'left', 'FontSize', 11, ...
        'ForegroundColor', [0.3, 0.3, 0.3]);

    % Define plate dimensions
    rows = 8;     % A-H
    cols = 12;    % 1-12
    wellSize = 49;
    wellSpacing = 8;
    startX = 30;
    startY = 30;
    
    % Create buttons for each well
    wellButtons = zeros(rows, cols);
    wellGroupAssignment = zeros(rows, cols);  % 0 = unassigned, 1+ = assigned to group N
    wellLabels = cell(rows, cols);
    
    % Calculate mapping between well positions and row labels
    for r = 1:rows
        for c = 1:cols
            idx = (r-1)*cols + c;
            if idx <= length(rowLabels)
                wellLabels{r, c} = rowLabels{idx};
            else
                wellLabels{r, c} = '';
            end
        end
    end
    
    % Column headers (1-12)
    for c = 1:cols
        uicontrol(platePanel, 'Style', 'text', 'String', num2str(c), ...
            'Position', [startX + (c-1)*(wellSize+wellSpacing), startY + rows*(wellSize+wellSpacing), wellSize, 20], ...
            'FontSize', 10, 'FontWeight', 'bold');
    end
    
    % Row headers (A-H)
    for r = 1:rows
        uicontrol(platePanel, 'Style', 'text', 'String', char(64+r), ...
            'Position', [startX - 25, startY + (rows-r)*(wellSize+wellSpacing), 20, wellSize], ...
            'FontSize', 10, 'FontWeight', 'bold');
    end
    
    % Create well buttons
    for r = 1:rows
        for c = 1:cols
            posX = startX + (c-1)*(wellSize+wellSpacing);
            posY = startY + (rows-r)*(wellSize+wellSpacing);
            
            wellButtons(r, c) = uicontrol(platePanel, 'Style', 'pushbutton', ...
                'String', sprintf('%s\n%s', char(64+r) + string(c), wellLabels{r, c}), ...
                'Position', [posX, posY, wellSize, wellSize], ...
                'FontSize', 7, ...
                'Callback', {@toggleWell, r, c});
            
            set(wellButtons(r, c), 'TooltipString', wellLabels{r, c});
        end
    end
    
    % Group information display in scrollable panel
    groupNameFields = cell(1, numGroups);
    groupStatusTexts = cell(1, numGroups);
    groupNavButtons = cell(1, numGroups);
    groupCheckboxes = cell(1, numGroups);  % NEW: Checkboxes for plot inclusion
    
    % Default group names
    groupNames = cell(1, numGroups);
    for i = 1:numGroups
        groupNames{i} = sprintf('Group %d', i);
    end
    
    % Calculate scroll panel dimensions
    elementHeight = 80;
    totalContentHeight = numGroups * elementHeight;
    
    % Calculate group panel height in pixels
    groupPanelPos = get(groupPanel, 'Position'); % Relative position
    figPos = get(wellFig, 'Position'); % Figure position in pixels
    groupPanelHeightPx = groupPanelPos(4) * figPos(4); % Panel height in pixels

    % Set scroll panel height - start at top and make it scrollable if content is larger than panel
    if totalContentHeight > groupPanelHeightPx
        % Content is larger than panel - make it scrollable
        % Position is [left, bottom, width, height]
        % Height should be relative to its parent (groupPanel), so it's > 1 if content is larger
        scrollPanelRelativeHeight = totalContentHeight / groupPanelHeightPx;
        set(scrollPanel, 'Position', [0.02, 1 - scrollPanelRelativeHeight + 0.02, 0.96, scrollPanelRelativeHeight]);
    else
        % Content fits in panel
        set(scrollPanel, 'Position', [0.02, 0.02, 0.96, 0.96]);
    end
    
    % Create scrollable group interface
    groupStartY = totalContentHeight - elementHeight + 50; % Position from top of scrollPanel
    
    for g = 1:numGroups
        yPos = groupStartY - (g-1) * elementHeight;
        
        % Color indicator
        uicontrol(scrollPanel, 'Style', 'frame', ...
            'Position', [10, yPos + 35, 20, 20], ...
            'BackgroundColor', groupColors{g});
        
        % Group name field
        uicontrol(scrollPanel, 'Style', 'text', 'String', 'Name:', ...
            'Position', [35, yPos + 35, 40, 20], 'HorizontalAlignment', 'left');
        
        groupNameFields{g} = uicontrol(scrollPanel, 'Style', 'edit', ...
            'String', groupNames{g}, ...
            'Position', [80, yPos + 35, 120, 20], ...
            'Callback', {@updateGroupNameCallback, g});
        
        % Group status
        groupStatusTexts{g} = uicontrol(scrollPanel, 'Style', 'text', ...
            'String', sprintf('Selected: 0/%d', repetitionsPerGroup), ...
            'Position', [10, yPos + 15, 120, 15], ...
            'HorizontalAlignment', 'left', 'FontSize', 9);
        
        % NEW: Checkbox for plot inclusion
        groupCheckboxes{g} = uicontrol(scrollPanel, 'Style', 'checkbox', ...
            'String', 'Include in plot', ...
            'Position', [140, yPos + 15, 100, 15], ...
            'FontSize', 9, 'Value', 1);  % Default checked
        
        % Navigation button
        groupNavButtons{g} = uicontrol(scrollPanel, 'Style', 'pushbutton', ...
            'String', sprintf('Select Group %d', g), ...
            'Position', [210, yPos + 35, 100, 25], ...
            'FontSize', 9, ...
            'Callback', {@selectGroupCallback, g});
    end
    
    % Add scroll functionality if needed
    if totalContentHeight > groupPanelHeightPx
        % Set scroll callback on the main figure
        set(wellFig, 'WindowScrollWheelFcn', {@scrollGroupPanel});
    end
    
   % Control buttons - FIXED POSITIONING
    clearGroupBtn = uicontrol(wellFig, 'Style', 'pushbutton', 'String', 'Clear Current Group', ...
        'Position', [50, 120, 120, 40], 'Callback', @clearCurrentGroupCallback);
    
    prevGroupBtn = uicontrol(wellFig, 'Style', 'pushbutton', 'String', '← Previous Group', ...
        'Position', [180, 120, 120, 40], 'Callback', @prevGroupCallback);
    
    nextGroupBtn = uicontrol(wellFig, 'Style', 'pushbutton', 'String', 'Next Group →', ...
        'Position', [310, 120, 120, 40], 'Callback', @nextGroupCallback);
    
    % Plot button (window stays open)
    plotBtn = uicontrol(wellFig, 'Style', 'pushbutton', 'String', 'Plot Results', ...
        'Position', [440, 120, 120, 40], 'FontWeight', 'bold', ...
        'BackgroundColor', [0.2, 0.8, 0.2], 'ForegroundColor', 'white', ...
        'Callback', @plotResultsCallback);
    
    
    % NEW: Enhanced Analysis button
    enhancedBtn = uicontrol(wellFig, 'Style', 'pushbutton', 'String', 'Enhanced Analysis', ...
        'Position', [570, 120, 130, 40], 'FontWeight', 'bold', ...
        'BackgroundColor', [0.2, 0.2, 0.8], 'ForegroundColor', 'white', ...
        'Enable', 'off', 'Callback', @enhancedAnalysisCallback);
    
    cancelBtn = uicontrol(wellFig, 'Style', 'pushbutton', 'String', 'Close Window', ...
        'Position', [710, 120, 100, 40], 'Callback', @cancelCallback);
    
    % Status text with progress
    statusText = uicontrol(wellFig, 'Style', 'text', ...
        'String', sprintf('Group 1: 0/%d wells selected (Progress: 0/%d groups completed)', ...
                         repetitionsPerGroup, numGroups), ...
        'Position', [50, 170, 700, 30], 'HorizontalAlignment', 'left', 'FontSize', 11);
    
    % Store data
    wellData = struct();
    wellData.wellGroupAssignment = wellGroupAssignment;
    wellData.wellLabels = wellLabels;
    wellData.rowLabels = rowLabels;
    wellData.currentGroup = 1;
    wellData.numGroups = numGroups;
    wellData.repetitionsPerGroup = repetitionsPerGroup;
    wellData.groupColors = groupColors;
    wellData.groupNames = groupNames;
    wellData.statusText = statusText; % Main status for well selection
    wellData.infoText = infoText; % Top info text
    wellData.wellStatusText = wellStatusText; % Dedicated status for well selection GUI
    wellData.groupNameFields = groupNameFields;
    wellData.groupStatusTexts = groupStatusTexts;
    wellData.groupNavButtons = groupNavButtons;
    wellData.groupCheckboxes = groupCheckboxes;  % NEW: Store checkboxes
    wellData.wellButtons = wellButtons;
    wellData.prevGroupBtn = prevGroupBtn;
    wellData.nextGroupBtn = nextGroupBtn;
    wellData.plotBtn = plotBtn;
    wellData.enhancedBtn = enhancedBtn;
    wellData.rows = rows;
    wellData.cols = cols;
    wellData.result = cell(1, numGroups);
    wellData.canceled = false;
    wellData.windowOpen = true;  % NEW: Track if window should stay open
    wellData.scrollPanel = scrollPanel;  % Store scroll panel reference
    wellData.scrollOffset = 0;  % Track scroll position
    
    setappdata(wellFig, 'wellData', wellData);
    
    % Initialize the UI
    updateStatusAndButtons();
    highlightCurrentGroupWells();
    
    % Store figure handle globally for external access
    setappdata(0, 'persistentWellSelectionFig', wellFig);
    
    % Wait for user action but don't block
    selectedRows = []; % Will be populated when Plot is clicked
    
    %% Nested callback functions
    
    % NEW: Scroll callback function
    function scrollGroupPanel(src, event)
        % Get current mouse position
        currentPoint = get(wellFig, 'CurrentPoint');
        
        % Get group panel position in figure coordinates
        groupPanelPos = get(groupPanel, 'Position');
        figPos = get(wellFig, 'Position');
        
        % Convert panel position to pixels
        panelLeft = groupPanelPos(1) * figPos(3);
        panelBottom = groupPanelPos(2) * figPos(4);
        panelWidth = groupPanelPos(3) * figPos(3);
        panelHeight = groupPanelPos(4) * figPos(4);
        
        % Check if mouse is over the group panel
        if currentPoint(1) >= panelLeft && currentPoint(1) <= (panelLeft + panelWidth) && ...
           currentPoint(2) >= panelBottom && currentPoint(2) <= (panelBottom + panelHeight)
            
            wellData = getappdata(wellFig, 'wellData');
            scrollStep = 50; % Pixels to scroll per step
            
            % Get current position
            currentPos = get(wellData.scrollPanel, 'Position');
            
            % Calculate scroll limits
            % Max scroll down: content bottom aligns with panel bottom
            % Max scroll up: content top aligns with panel top (currentPos(2) = 1 - scrollPanelRelativeHeight)
            
            scrollPanelRelativeHeight = get(wellData.scrollPanel, 'Position');
            scrollPanelRelativeHeight = scrollPanelRelativeHeight(4); % Get the actual relative height
            
            maxScrollDown = 1 - scrollPanelRelativeHeight + 0.02; % Bottom of content aligns with bottom of panel
            maxScrollUp = 0.02; % Top of content aligns with top of panel (approx)
            
            % Calculate new position
            if event.VerticalScrollCount > 0 % Scroll up (mouse wheel down moves content up)
                newY = min(currentPos(2) + scrollStep / groupPanelHeightPx, maxScrollUp);
            else % Scroll down (mouse wheel up moves content down)
                newY = max(currentPos(2) - scrollStep / groupPanelHeightPx, maxScrollDown);
            end
            
            % Apply new position
            set(wellData.scrollPanel, 'Position', [currentPos(1), newY, currentPos(3), currentPos(4)]);
            setappdata(wellFig, 'wellData', wellData);
        end
    end
    
    function toggleWell(src, ~, row, col)
        wellData = getappdata(wellFig, 'wellData');
        currentGroup = wellData.currentGroup;
        
        % Get current assignment
        currentAssignment = wellData.wellGroupAssignment(row, col);
        
        % Toggle well group assignment
        if currentAssignment == currentGroup
            % Unassign
            wellData.wellGroupAssignment(row, col) = 0;
            set(src, 'BackgroundColor', [0.94, 0.94, 0.94]);
            set(src, 'FontWeight', 'normal');
        elseif currentAssignment == 0
            % Assign to current group
            wellData.wellGroupAssignment(row, col) = currentGroup;
            set(src, 'BackgroundColor', wellData.groupColors{currentGroup});
            set(src, 'FontWeight', 'bold');
            
            % Check if we've reached the expected number of repetitions
            currentCount = sum(wellData.wellGroupAssignment(:) == currentGroup);
            if currentCount >= wellData.repetitionsPerGroup && currentGroup < wellData.numGroups
                % Auto-advance to next group
                wellData.currentGroup = currentGroup + 1;
                setappdata(wellFig, 'wellData', wellData);
                updateStatusAndButtons();
                highlightCurrentGroupWells();
                return;
            end
        else
            % Already assigned to another group
            set(wellData.wellStatusText, 'String', sprintf('Well %s%d is already assigned to Group %d. Clear it first if you want to reassign.', ...
                char(64+row), col, currentAssignment), 'ForegroundColor', [0.8, 0.2, 0.2]);
            return;
        end
        
        setappdata(wellFig, 'wellData', wellData);
        updateStatusAndButtons();
        set(wellData.wellStatusText, 'String', 'Ready.', 'ForegroundColor', [0.3, 0.3, 0.3]); % Clear previous message
    end
    
    function selectGroupCallback(~, ~, groupNumber)
        wellData = getappdata(wellFig, 'wellData');
        wellData.currentGroup = groupNumber;
        setappdata(wellFig, 'wellData', wellData);
        updateStatusAndButtons();
        highlightCurrentGroupWells();
    end
    
    function prevGroupCallback(~, ~)
        wellData = getappdata(wellFig, 'wellData');
        if wellData.currentGroup > 1
            wellData.currentGroup = wellData.currentGroup - 1;
            setappdata(wellFig, 'wellData', wellData);
            updateStatusAndButtons();
            highlightCurrentGroupWells();
        end
    end
    
    function nextGroupCallback(~, ~)
        wellData = getappdata(wellFig, 'wellData');
        if wellData.currentGroup < wellData.numGroups
            wellData.currentGroup = wellData.currentGroup + 1;
            setappdata(wellFig, 'wellData', wellData);
            updateStatusAndButtons();
            highlightCurrentGroupWells();
        end
    end
    
    function updateGroupNameCallback(src, ~, groupNumber)
        wellData = getappdata(wellFig, 'wellData');
        newName = get(src, 'String');
        wellData.groupNames{groupNumber} = newName;
        setappdata(wellFig, 'wellData', wellData);
        updateStatusAndButtons();
    end
    
    function clearCurrentGroupCallback(~, ~)
        wellData = getappdata(wellFig, 'wellData');
        currentGroup = wellData.currentGroup;
        
        % Clear all wells assigned to current group
        for r = 1:wellData.rows
            for c = 1:wellData.cols
                if wellData.wellGroupAssignment(r, c) == currentGroup
                    wellData.wellGroupAssignment(r, c) = 0;
                    set(wellData.wellButtons(r, c), 'BackgroundColor', [0.94, 0.94, 0.94]);
                    set(wellData.wellButtons(r, c), 'FontWeight', 'normal');
                end
            end
        end
        
        setappdata(wellFig, 'wellData', wellData);
        updateStatusAndButtons();
    end
    
    % MODIFIED: Plot results callback (doesn't close window)
    function plotResultsCallback(~, ~)
        wellData = getappdata(wellFig, 'wellData');
        
        % Get final group names and check which groups are selected for plotting
        finalGroupNames = cell(1, wellData.numGroups);
        groupsToPlot = false(1, wellData.numGroups);
        
        for i = 1:wellData.numGroups
            finalGroupNames{i} = get(wellData.groupNameFields{i}, 'String');
            if isempty(finalGroupNames{i})
                finalGroupNames{i} = sprintf('Group %d', i);
            end
            % NEW: Check if group is marked for plotting
            groupsToPlot(i) = get(wellData.groupCheckboxes{i}, 'Value');
        end
        
        % Process results for each group
        allGroupRows = cell(1, wellData.numGroups);
        for g = 1:wellData.numGroups
            groupWells = (wellData.wellGroupAssignment == g);
            selectedIndices = find(groupWells);
            selectedLabelsForGroup = cell(1, length(selectedIndices));
            
            % Get labels for selected wells
            for i = 1:length(selectedIndices)
                [r, c] = ind2sub([wellData.rows, wellData.cols], selectedIndices(i));
                selectedLabelsForGroup{i} = wellData.wellLabels{r, c};
            end
            
            % Map back to original indices
            selectedRowsForGroup = [];
            for i = 1:length(selectedLabelsForGroup)
                for j = 1:length(wellData.rowLabels)
                    if strcmp(selectedLabelsForGroup{i}, wellData.rowLabels{j})
                        selectedRowsForGroup = [selectedRowsForGroup, j];
                        break;
                    end
                end
            end
            
            allGroupRows{g} = selectedRowsForGroup;
        end
        
        % NEW: Filter groups based on checkboxes
        filteredGroupRows = cell(1, sum(groupsToPlot));
        filteredGroupNames = cell(1, sum(groupsToPlot));
        filterIndex = 1;
        
        for g = 1:wellData.numGroups
            if groupsToPlot(g) && ~isempty(allGroupRows{g})
                filteredGroupRows{filterIndex} = allGroupRows{g};
                filteredGroupNames{filterIndex} = finalGroupNames{g};
                filterIndex = filterIndex + 1;
            end
        end
        
        % Check if any groups are selected for plotting
        if isempty(filteredGroupRows) || all(cellfun(@isempty, filteredGroupRows))
            set(wellData.wellStatusText, 'String', 'No groups selected for plotting or no wells assigned. Please check group selections and checkboxes.', ...
                'ForegroundColor', [0.8, 0.2, 0.2]);
            return;
        end
        
        % Store results for external access
        result = struct();
        result.selectedRows = filteredGroupRows;
        result.groupNames = filteredGroupNames;
        result.windowOpen = true;
        
        wellData.result = result;
        setappdata(wellFig, 'wellData', wellData);
        
        % Store globally for external access
        setappdata(wellFig, 'latestResult', result);
        
        % Update status to show plot was generated
        plotInfo = sprintf('Plot generated with %d groups: %s', length(filteredGroupNames), strjoin(filteredGroupNames, ', '));
        set(wellData.wellStatusText, 'String', plotInfo);
        set(wellData.wellStatusText, 'ForegroundColor', [0, 0.6, 0]);

         % Update status to show plot was generated
        plotInfo = sprintf('Plot generated with %d groups: %s', length(filteredGroupNames), strjoin(filteredGroupNames, ', '));
        set(wellData.statusText, 'String', plotInfo);
        set(wellData.statusText, 'ForegroundColor', [0, 0.6, 0]);
        
        % NEW: Enable enhanced analysis button
        set(enhancedBtn, 'Enable', 'on');
    end

 % NEW: Enhanced Analysis callback
    function enhancedAnalysisCallback(~, ~)
        enhancedData = getappdata(0, 'enhancedAnalysisData');
        
        if isempty(enhancedData)
            set(wellData.wellStatusText, 'String', 'No analysis data found. Please run basic analysis first by clicking "Plot Results".', ...
                'ForegroundColor', [0.8, 0.2, 0.2]);
            return;
        end
        
        % Launch enhanced analysis window
        growth_analysis_module.launchEnhancedAnalysisGUI(enhancedData.groupData, enhancedData.groupAverage, ...
                                 enhancedData.groupStd, enhancedData.groupNames, ...
                                 enhancedData.timeVectorInHours, enhancedData.dataType, ...
                                 enhancedData.outputFolderPath, enhancedData.timestamp, ...
                                 enhancedData.textData, enhancedData.groupRows);
    end
    
    % NEW: Close window callback
    function cancelCallback(~, ~)
        wellData = getappdata(wellFig, 'wellData');
        wellData.canceled = true;
        wellData.windowOpen = false;
        setappdata(wellFig, 'wellData', wellData);
        delete(wellFig);
    end
    
    function updateStatusAndButtons()
        wellData = getappdata(wellFig, 'wellData');
        
        % Count wells for each group
        groupCounts = zeros(1, wellData.numGroups);
        for g = 1:wellData.numGroups
            groupCounts(g) = sum(wellData.wellGroupAssignment(:) == g);
        end
        
        % Update group status texts
        for g = 1:wellData.numGroups
            count = groupCounts(g);
            expected = wellData.repetitionsPerGroup;
            if count >= expected
                status_str = sprintf('✓ Selected: %d/%d', count, expected);
                color = [0, 0.6, 0];
            elseif count > 0
                status_str = sprintf('Selected: %d/%d', count, expected);
                color = [0.8, 0.5, 0];
            else
                status_str = sprintf('Selected: %d/%d', count, expected);
                color = [0.5, 0.5, 0.5];
            end
            set(wellData.groupStatusTexts{g}, 'String', status_str, 'ForegroundColor', color);
        end
        
        % Update navigation buttons
        if wellData.currentGroup > 1
            set(wellData.prevGroupBtn, 'Enable', 'on');
        else
            set(wellData.prevGroupBtn, 'Enable', 'off');
        end
        
        if wellData.currentGroup < wellData.numGroups
            set(wellData.nextGroupBtn, 'Enable', 'on');
        else
            set(wellData.nextGroupBtn, 'Enable', 'off');
        end
        
        % Update main status
        currentCount = groupCounts(wellData.currentGroup);
        completedGroups = sum(groupCounts >= wellData.repetitionsPerGroup);
        
        currentGroupName = get(wellData.groupNameFields{wellData.currentGroup}, 'String');
        
        set(wellData.statusText, 'String', sprintf('%s: %d/%d wells selected (Progress: %d/%d groups completed)', ...
            currentGroupName, currentCount, wellData.repetitionsPerGroup, completedGroups, wellData.numGroups));
        
        set(wellData.infoText, 'String', sprintf('Select %d wells for %s (%d of %d groups)', ...
            wellData.repetitionsPerGroup, currentGroupName, wellData.currentGroup, wellData.numGroups));
        
        % Highlight current group navigation button
        for g = 1:wellData.numGroups
            if g == wellData.currentGroup
                set(wellData.groupNavButtons{g}, 'BackgroundColor', [0.3, 0.6, 1]);
                set(wellData.groupNavButtons{g}, 'FontWeight', 'bold');
            else
                set(wellData.groupNavButtons{g}, 'BackgroundColor', [0.94, 0.94, 0.94]);
                set(wellData.groupNavButtons{g}, 'FontWeight', 'normal');
            end
        end
        
        % Enable plot button if at least one group has selections and is checked
        anyGroupReady = false;
        for g = 1:wellData.numGroups
            if groupCounts(g) > 0 && get(wellData.groupCheckboxes{g}, 'Value')
                anyGroupReady = true;
                break;
            end
        end
        
        if anyGroupReady
            set(wellData.plotBtn, 'Enable', 'on');
        else
            set(wellData.plotBtn, 'Enable', 'off');
        end
    end
    
    function highlightCurrentGroupWells()
        wellData = getappdata(wellFig, 'wellData');
        currentGroup = wellData.currentGroup;
        
        % Update all buttons to show appropriate state
        for r = 1:wellData.rows
            for c = 1:wellData.cols
                currentAssignment = wellData.wellGroupAssignment(r, c);
                
                if currentAssignment == currentGroup
                    % Well belongs to current group
                    set(wellData.wellButtons(r, c), 'BackgroundColor', wellData.groupColors{currentGroup});
                    set(wellData.wellButtons(r, c), 'FontWeight', 'bold');
                elseif currentAssignment > 0
                    % Well belongs to different group - show dimmed
                    otherGroup = currentAssignment;
                    baseColor = wellData.groupColors{otherGroup};
                    dimmedColor = baseColor * 0.6 + [0.4, 0.4, 0.4];
                    set(wellData.wellButtons(r, c), 'BackgroundColor', dimmedColor);
                    set(wellData.wellButtons(r, c), 'FontWeight', 'normal');
                else
                    % Unassigned well
                    set(wellData.wellButtons(r, c), 'BackgroundColor', [0.94, 0.94, 0.94]);
                    set(wellData.wellButtons(r, c), 'FontWeight', 'normal');
                end
            end
        end
    end
        end
        
        function output = convertToNumeric(input)
if isempty(input) || ischar(input) || iscell(input)
        output = 0;
    else
        output = double(input);
end
        end
        
        function result = safeDiv(a, b)
% Safe division that handles edge cases
    if isempty(a) || isempty(b)
        result = 0;
    else
        a = double(a);
        b = double(b);
        if b == 0
            result = 0;
        else
            result = a / b;
            if result < 0
                result = 0;
            end
        end
    end
        end
        
        function output = convertMissing(input)
if ismissing(input)
        output = '';
    else
        output = input;
end
        end
        
        function analyzeLumVsODGUI(mainFig)
handles = guidata(mainFig);
    set(handles.statusText, 'String', 'Step 4: Selecting matrices for Lum vs. OD analysis...');
    
    % Select first file (Y-axis data)
    [filename1, filepath1] = uigetfile({'*.xlsx;*.xls', 'Excel Files (*.xlsx, *.xls)'}, ...
                            'Select First Matrix File (Y-axis data)', 'LuciferaseAN.xlsx');
    if isequal(filename1, 0)
        set(handles.statusText, 'String', 'No file selected. Ready for next operation.');
        return;
    end
    
    % Select second file (X-axis data)
    [filename2, filepath2] = uigetfile({'*.xlsx;*.xls', 'Excel Files (*.xlsx, *.xls)'}, ...
                            'Select Second Matrix File (X-axis data)', 'ODgrowthAN.xlsx');
    if isequal(filename2, 0)
        set(handles.statusText, 'String', 'No file selected. Ready for next operation.');
        return;
    end
    
    try
        % Read both matrices
        [dataMatrix1, textData1] = xlsread(fullfile(filepath1, filename1));
        [dataMatrix2, textData2] = xlsread(fullfile(filepath2, filename2));
        
        % Extract time vectors (assuming first row contains time)
        timeVector1 = dataMatrix1(1, :);
        timeVector2 = dataMatrix2(1, :);
        
        % Remove time row from data
        dataMatrix1 = dataMatrix1(2:end, :);
        dataMatrix2 = dataMatrix2(2:end, :);
        textData1 = textData1(2:end, 1);
        textData2 = textData2(2:end, 1);
        
        % Get number of groups
        prompt = {'Enter the number of groups (e.g., 1):'};
        dlgtitle = 'Number of Groups';
        dims = [1 50];
        definput = {'1'};
        numGroupsStr = inputdlg(prompt, dlgtitle, dims, definput);
        
        if isempty(numGroupsStr)
            set(handles.statusText, 'String', 'Operation canceled. Ready for next operation.');
            return;
        end
        
        numGroups = str2double(numGroupsStr{1});
        % Input validation for numGroups
        if isnan(numGroups) || numGroups < 1 || floor(numGroups) ~= numGroups
            errordlg('Number of groups must be a positive integer.', 'Invalid Input');
            set(handles.statusText, 'String', 'Invalid input. Operation canceled.');
            return;
        end

        % Get number of repetitions per group
        prompt = {'Enter the number of repetitions per group (e.g., 3):'};
        dlgtitle = 'Repetitions per Group';
        dims = [1 50];
        definput = {'3'};
        repetitionsStr = inputdlg(prompt, dlgtitle, dims, definput);
        
        if isempty(repetitionsStr)
            set(handles.statusText, 'String', 'Operation canceled. Ready for next operation.');
            return;
        end
        
        repetitionsPerGroup = str2double(repetitionsStr{1});
        % Input validation for repetitionsPerGroup
        if isnan(repetitionsPerGroup) || repetitionsPerGroup < 1 || floor(repetitionsPerGroup) ~= repetitionsPerGroup
            errordlg('Number of repetitions per group must be a positive integer.', 'Invalid Input');
            set(handles.statusText, 'String', 'Invalid input. Operation canceled.');
            return;
        end

        % Generate colors for groups
        groupColors = cell(1, numGroups);
        for i = 1:numGroups
            groupColors{i} = growth_analysis_module.generateUniqueColor(i);
        end
        
        % Create the persistent well selection window
        set(handles.statusText, 'String', 'Opening well selection window for Lum vs OD analysis...');
        drawnow;
        
        % Call the enhanced well selection GUI (it will stay open)
        growth_analysis_module.selectWellsEnhancedGUI(textData1, numGroups, repetitionsPerGroup);
        
        % Store analysis data for Step 4
        analysisData = struct();
        analysisData.dataMatrix1 = dataMatrix1;  % Luciferase data
        analysisData.dataMatrix2 = dataMatrix2;  % OD data
        analysisData.textData1 = textData1;
        analysisData.textData2 = textData2;
        analysisData.timeVector1 = timeVector1;
        analysisData.timeVector2 = timeVector2;
        analysisData.filepath1 = filepath1;
        analysisData.filepath2 = filepath2;
        analysisData.groupColors = groupColors;
        analysisData.handles = handles;
        analysisData.analysisType = 'LumVsOD';
        
        % Store analysis data globally
        setappdata(0, 'currentAnalysisData', analysisData);
        
        % Set up a timer to check for plot requests from the well selection window
        if isempty(timerfind('Name', 'WellSelectionMonitor'))
            t = timer('Name', 'WellSelectionMonitor', ...
                     'Period', 0.5, ...
                     'ExecutionMode', 'fixedRate', ...
                     'TimerFcn', @growth_analysis_module.checkForLumVsODPlotRequest);
            start(t);
        end
        
        % Update main status
        set(handles.statusText, 'String', [
            'Step 4: Well selection window opened!' newline newline ...
            'Instructions:' newline ...
            '1. Select wells for each group in the selection window' newline ...
            '2. Check/uncheck groups to include in plot' newline ...
            '3. Click "Plot Results" in the selection window' newline ...
            '4. Window stays open for multiple plots' newline newline ...
            'Ready for Lum vs OD analysis!']);
        
    catch ME
        set(handles.statusText, 'String', ['Error during Lum vs. OD analysis setup: ' ME.message newline newline ...
                                          'Please check your input files and try again.']);
        disp(getReport(ME, 'extended'));
    end
        end
        
        function analyze3DLumVsODGUI(mainFig)
handles = guidata(mainFig);
    set(handles.statusText, 'String', 'Step 5: Selecting matrices for 3D Lum vs. OD vs. Time analysis...');
    
    % Select first file (Y-axis data)
    [filename1, filepath1] = uigetfile({'*.xlsx;*.xls', 'Excel Files (*.xlsx, *.xls)'}, ...
                            'Select First Matrix File (Y-axis data)', 'LuciferaseAN.xlsx');
    if isequal(filename1, 0)
        set(handles.statusText, 'String', 'No file selected. Ready for next operation.');
        return;
    end
    
    % Select second file (X-axis data)
    [filename2, filepath2] = uigetfile({'*.xlsx;*.xls', 'Excel Files (*.xlsx, *.xls)'}, ...
                            'Select Second Matrix File (X-axis data)', 'ODgrowthAN.xlsx');
    if isequal(filename2, 0)
        set(handles.statusText, 'String', 'No file selected. Ready for next operation.');
        return;
    end
    
    try
        % Read both matrices
        [dataMatrix1, textData1] = xlsread(fullfile(filepath1, filename1));
        [dataMatrix2, textData2] = xlsread(fullfile(filepath2, filename2));
        
        % Extract time vectors (assuming first row contains time)
        timeVector1 = dataMatrix1(1, :);
        timeVector2 = dataMatrix2(1, :);
        timeVectorHours = timeVector1 / 3600; % Convert to hours
        
        % Remove time row from data
        dataMatrix1 = dataMatrix1(2:end, :);
        dataMatrix2 = dataMatrix2(2:end, :);
        textData1 = textData1(2:end, 1);
        textData2 = textData2(2:end, 1);
        
        % Get number of groups
        prompt = {'Enter the number of groups (e.g., 1):'};
        dlgtitle = 'Number of Groups';
        dims = [1 50];
        definput = {'1'};
        numGroupsStr = inputdlg(prompt, dlgtitle, dims, definput);
        
        if isempty(numGroupsStr)
            set(handles.statusText, 'String', 'Operation canceled. Ready for next operation.');
            return;
        end
        
        numGroups = str2double(numGroupsStr{1});
        % Input validation for numGroups
        if isnan(numGroups) || numGroups < 1 || floor(numGroups) ~= numGroups
            errordlg('Number of groups must be a positive integer.', 'Invalid Input');
            set(handles.statusText, 'String', 'Invalid input. Operation canceled.');
            return;
        end

        % Get number of repetitions per group
        prompt = {'Enter the number of repetitions per group (e.g., 3):'};
        dlgtitle = 'Repetitions per Group';
        dims = [1 50];
        definput = {'3'};
        repetitionsStr = inputdlg(prompt, dlgtitle, dims, definput);
        
        if isempty(repetitionsStr)
            set(handles.statusText, 'String', 'Operation canceled. Ready for next operation.');
            return;
        end
        
        repetitionsPerGroup = str2double(repetitionsStr{1});
        % Input validation for repetitionsPerGroup
        if isnan(repetitionsPerGroup) || repetitionsPerGroup < 1 || floor(repetitionsPerGroup) ~= repetitionsPerGroup
            errordlg('Number of repetitions per group must be a positive integer.', 'Invalid Input');
            set(handles.statusText, 'String', 'Invalid input. Operation canceled.');
            return;
        end

        % Generate colors for groups
        groupColors = cell(1, numGroups);
        for i = 1:numGroups
            groupColors{i} = growth_analysis_module.generateUniqueColor(i);
        end
        
        % Create the persistent well selection window
        set(handles.statusText, 'String', 'Opening well selection window for 3D analysis...');
        drawnow;
        
        % Call the enhanced well selection GUI (it will stay open)
        growth_analysis_module.selectWellsEnhancedGUI(textData1, numGroups, repetitionsPerGroup);
        
        % Store analysis data for Step 5
        analysisData = struct();
        analysisData.dataMatrix1 = dataMatrix1;  % Luciferase data
        analysisData.dataMatrix2 = dataMatrix2;  % OD data
        analysisData.textData1 = textData1;
        analysisData.textData2 = textData2;
        analysisData.timeVector1 = timeVector1;
        analysisData.timeVector2 = timeVector2;
        analysisData.timeVectorHours = timeVectorHours;
        analysisData.filepath1 = filepath1;
        analysisData.filepath2 = filepath2;
        analysisData.groupColors = groupColors;
        analysisData.handles = handles;
        analysisData.analysisType = '3DLumVsOD';
        
        % Store analysis data globally
        setappdata(0, 'currentAnalysisData', analysisData);
        
        % Set up a timer to check for plot requests from the well selection window
        if isempty(timerfind('Name', 'WellSelectionMonitor'))
            t = timer('Name', 'WellSelectionMonitor', ...
                     'Period', 0.5, ...
                     'ExecutionMode', 'fixedRate', ...
                     'TimerFcn', @growth_analysis_module.checkFor3DPlotRequest);
            start(t);
        end
        
        % Update main status
        set(handles.statusText, 'String', [
            'Step 5: Well selection window opened!' newline newline ...
            'Instructions:' newline ...
            '1. Select wells for each group in the selection window' newline ...
            '2. Check/uncheck groups to include in plot' newline ...
            '3. Click "Plot Results" in the selection window' newline ...
            '4. Window stays open for multiple plots' newline newline ...
            'Ready for 3D Lum vs OD vs Time analysis!']);
        
    catch ME
        set(handles.statusText, 'String', ['Error during 3D analysis setup: ' ME.message newline newline ...
                                          'Please check your input files and try again.']);
        disp(getReport(ME, 'extended'));
    end
        end
        
        function checkForLumVsODPlotRequest(~, ~)
% Check if there's a plot request from the well selection window for Lum vs OD analysis
    
    figHandle = getappdata(0, 'persistentWellSelectionFig');
    
    if isempty(figHandle) || ~ishandle(figHandle)
        % Well selection window is closed, stop timer
        t = timerfind('Name', 'WellSelectionMonitor');
        if ~isempty(t)
            stop(t);
            delete(t);
        end
        return;
    end
    
    % Check if there's a new result
    result = getappdata(figHandle, 'latestResult');
    
    if ~isempty(result)
        % Get current time from well selection window data
        wellData = getappdata(figHandle, 'wellData');
        if isfield(wellData, 'statusText') && ishandle(wellData.statusText)
            statusString = get(wellData.statusText, 'String');
            
            % Simple check: if status contains "Plot generated", process it
            if contains(statusString, 'Plot generated') && ~contains(statusString, 'processed')
                growth_analysis_module.performLumVsODAnalysis(result);
                
                % Mark as processed
                set(wellData.statusText, 'String', [statusString ' (processed)']);
            end
        end
    end
        end
        
        function checkFor3DPlotRequest(~, ~)
% Check if there's a plot request from the well selection window for 3D analysis
    
    figHandle = getappdata(0, 'persistentWellSelectionFig');
    
    if isempty(figHandle) || ~ishandle(figHandle)
        % Well selection window is closed, stop timer
        t = timerfind('Name', 'WellSelectionMonitor');
        if ~isempty(t)
            stop(t);
            delete(t);
        end
        return;
    end
    
    % Check if there's a new result
    result = getappdata(figHandle, 'latestResult');
    
    if ~isempty(result)
        % Get current time from well selection window data
        wellData = getappdata(figHandle, 'wellData');
        if isfield(wellData, 'statusText') && ishandle(wellData.statusText)
            statusString = get(wellData.statusText, 'String');
            
            % Simple check: if status contains "Plot generated", process it
            if contains(statusString, 'Plot generated') && ~contains(statusString, 'processed')
                growth_analysis_module.perform3DAnalysis(result);
                
                % Mark as processed
                set(wellData.statusText, 'String', [statusString ' (processed)']);
            end
        end
    end
        end
        
        function performLumVsODAnalysis(result)
% Perform Lum vs OD analysis with the given result
    
    % Get analysis data
    analysisData = getappdata(0, 'currentAnalysisData');
    if isempty(analysisData) || ~strcmp(analysisData.analysisType, 'LumVsOD')
        return;
    end
    
    try
        % Extract data
        dataMatrix1 = analysisData.dataMatrix1;  % Luciferase
        dataMatrix2 = analysisData.dataMatrix2;  % OD
        textData1 = analysisData.textData1;
        groupColors = analysisData.groupColors;
        filepath1 = analysisData.filepath1;
        handles = analysisData.handles;
        
        % Extract group information from result
        groupNames = result.groupNames;
        groupRows = result.selectedRows;
        numGroups = length(groupNames);
        
        % Create output folder
        outputFolderPath = fullfile(filepath1, 'LumVsOD_graphs');
        if ~exist(outputFolderPath, 'dir')
            mkdir(outputFolderPath);
        end
        
        % Generate timestamp
        timestamp = datestr(now, 'yyyymmdd_HHMMSS');
        
        % Process each group and create scatter plots (colored by wells)
        for i = 1:numGroups
            if ~isempty(groupRows{i})
                % Get data for this group from both matrices
                yData = dataMatrix1(groupRows{i}, :);  % Luciferase data
                xData = dataMatrix2(groupRows{i}, :);  % OD data
                numWells = size(yData, 1);
                
                % Create scatter plot with colors by wells
                figGroup = figure('Name', sprintf('Group %d: %s_%s', i, groupNames{i}, timestamp));
                
                % Generate unique colors for each well
                wellColors = lines(numWells);
                
                hold on;
                for w = 1:numWells
                    wellName = textData1{groupRows{i}(w)};
                    scatter(xData(w, :), yData(w, :), 50, wellColors(w, :), 'filled', ...
                           'DisplayName', wellName);
                end
                
                xlabel('OD 600nm', 'FontSize', 12, 'FontWeight', 'bold');
                ylabel('Luciferase', 'FontSize', 12, 'FontWeight', 'bold');
                title([groupNames{i} ': Luciferase vs OD (colored by wells)'], 'FontSize', 14, 'FontWeight', 'bold');
                legend('Location', 'best', 'FontSize', 8);
                grid on;
                hold off;
                
                % Save the plot
                figFilename = sprintf('%s_LumVsOD_byWells_%s.png', groupNames{i}, timestamp);
                saveas(figGroup, fullfile(outputFolderPath, figFilename));
            end
        end
        
        % Create combined plot for all groups at final time point
        figCombined = figure('Name', sprintf('All Groups: Luciferase vs OD_%s', timestamp));
        hold on;
        
        for i = 1:numGroups
            if ~isempty(groupRows{i})
                % Use final time point
                finalTimeIdx = size(dataMatrix1, 2);
                yData = dataMatrix1(groupRows{i}, finalTimeIdx);
                xData = dataMatrix2(groupRows{i}, finalTimeIdx);
                
                color = groupColors{i};
                scatter(xData, yData, 100, color, 'filled', ...
                       'DisplayName', groupNames{i});
            end
        end
        
        xlabel('OD 600nm (Final)', 'FontSize', 12, 'FontWeight', 'bold');
        ylabel('Luciferase (Final)', 'FontSize', 12, 'FontWeight', 'bold');
        title('All Groups: Final Luciferase vs OD', 'FontSize', 14, 'FontWeight', 'bold');
        legend('Location', 'best');
        grid on;
        hold off;
        
        % Save combined plot
        combinedFilename = sprintf('All_Groups_LumVsOD_%s.png', timestamp);
        saveas(figCombined, fullfile(outputFolderPath, combinedFilename));
        
        % Update status
        set(handles.statusText, 'String', [
            'Lum vs. OD analysis completed successfully!' newline newline ...
            'Results saved to: ' outputFolderPath newline ...
            'Generated files:' newline ...
            '• Individual group scatter plots' newline ...
            '• Combined final time point plot' newline newline ...
            'Well selection window remains open for additional analyses.']);
        
    catch ME
        set(handles.statusText, 'String', ['Error during Lum vs. OD analysis: ' ME.message newline newline ...
                                          'Please check your input files and try again.']);
        disp(getReport(ME, 'extended'));
    end
        end
        
        function perform3DAnalysis(result)
% Perform 3D Lum vs OD vs Time analysis with the given result
    
    % Get analysis data
    analysisData = getappdata(0, 'currentAnalysisData');
    if isempty(analysisData) || ~strcmp(analysisData.analysisType, '3DLumVsOD')
        return;
    end
    
    try
        % Extract data
        dataMatrix1 = analysisData.dataMatrix1;  % Luciferase
        dataMatrix2 = analysisData.dataMatrix2;  % OD
        textData1 = analysisData.textData1;
        timeVectorHours = analysisData.timeVectorHours;
        groupColors = analysisData.groupColors;
        filepath1 = analysisData.filepath1;
        handles = analysisData.handles;
        
        % Extract group information from result
        groupNames = result.groupNames;
        groupRows = result.selectedRows;
        numGroups = length(groupNames);
        
        % Create output folder
        outputFolderPath = fullfile(filepath1, '3D_LumVsOD_graphs');
        if ~exist(outputFolderPath, 'dir')
            mkdir(outputFolderPath);
        end
        
        % Generate timestamp
        timestamp = datestr(now, 'yyyymmdd_HHMMSS');
        
        % Create individual 3D plots for each group
        for i = 1:numGroups
            if ~isempty(groupRows{i})
                % Get data for this group from both matrices
                yData = dataMatrix1(groupRows{i}, :);  % Luciferase data
                xData = dataMatrix2(groupRows{i}, :);  % OD data
                numWells = size(yData, 1);
                numTimePoints = size(yData, 2);
                
                % Create 3D plot
                fig3D = figure('Name', sprintf('3D: %s_%s', groupNames{i}, timestamp), 'Position', [100, 100, 800, 600]);
                
                % Generate colors for each well
                wellColors = lines(numWells);
                
                hold on;
                for w = 1:numWells
                    wellName = textData1{groupRows{i}(w)};
                    
                    % Create 3D trajectory for this well
                    plot3(xData(w, :), yData(w, :), timeVectorHours, ...
                          'Color', wellColors(w, :), 'LineWidth', 2, ...
                          'DisplayName', wellName);
                    
                    % Add markers at specific time points
                    markerIndices = 1:5:numTimePoints; % Every 5th time point
                    scatter3(xData(w, markerIndices), yData(w, markerIndices), ...
                            timeVectorHours(markerIndices), 50, wellColors(w, :), ...
                            'filled', 'MarkerEdgeColor', 'k');
                end
                
                xlabel('OD 600nm', 'FontSize', 12, 'FontWeight', 'bold');
                ylabel('Luciferase', 'FontSize', 12, 'FontWeight', 'bold');
                zlabel('Time (hours)', 'FontSize', 12, 'FontWeight', 'bold');
                title([groupNames{i} ': 3D Luciferase vs OD vs Time'], 'FontSize', 14, 'FontWeight', 'bold');
                legend('Location', 'best', 'FontSize', 8);
                grid on;
                view(45, 30); % Set a good viewing angle
                hold off;
                
                % Save the 3D plot
                figFilename = sprintf('%s_3D_LumVsOD_%s', groupNames{i}, timestamp);
                saveas(fig3D, fullfile(outputFolderPath, [figFilename '.png']));
                saveas(fig3D, fullfile(outputFolderPath, [figFilename '.fig'])); % Save as .fig for rotation
            end
        end
        
        % Create combined 3D plot for all groups
        figCombined3D = figure('Name', sprintf('All Groups: 3D Luciferase vs OD vs Time_%s', timestamp), ...
                              'Position', [200, 200, 1000, 800]);
        
        hold on;
        legendEntries = {};
        
        for i = 1:numGroups
            if ~isempty(groupRows{i})
                % Get data for this group
                yData = dataMatrix1(groupRows{i}, :);  % Luciferase data
                xData = dataMatrix2(groupRows{i}, :);  % OD data
                numWells = size(yData, 1);
                
                % Plot each well in this group with group color
                for w = 1:numWells
                    wellName = textData1{groupRows{i}(w)};
                    
                    % Create 3D trajectory
                    plot3(xData(w, :), yData(w, :), timeVectorHours, ...
                          'Color', groupColors{i}, 'LineWidth', 1.5, ...
                          'LineStyle', '-');
                    
                    % Add start and end markers
                    scatter3(xData(w, 1), yData(w, 1), timeVectorHours(1), ...
                            100, groupColors{i}, 'o', 'filled', ...
                            'MarkerEdgeColor', 'k', 'LineWidth', 1);
                    scatter3(xData(w, end), yData(w, end), timeVectorHours(end), ...
                            100, groupColors{i}, 's', 'filled', ...
                            'MarkerEdgeColor', 'k', 'LineWidth', 1);
                end
                
                % Add to legend (one entry per group)
                legendEntries{end+1} = [groupNames{i} ' (○ start, ■ end)'];
            end
        end
        
        xlabel('OD 600nm', 'FontSize', 12, 'FontWeight', 'bold');
        ylabel('Luciferase', 'FontSize', 12, 'FontWeight', 'bold');
        zlabel('Time (hours)', 'FontSize', 12, 'FontWeight', 'bold');
        title('All Groups: 3D Luciferase vs OD vs Time Trajectories', 'FontSize', 14, 'FontWeight', 'bold');
        
        % Create custom legend with group colors
        legendHandles = [];
        for i = 1:length(legendEntries)
            h = plot3(NaN, NaN, NaN, 'Color', groupColors{i}, 'LineWidth', 2);
            legendHandles = [legendHandles, h];
        end
        legend(legendHandles, legendEntries, 'Location', 'best', 'FontSize', 10);
        
        grid on;
        view(45, 30);
        hold off;
        
        % Save combined 3D plot
        combinedFilename = sprintf('All_Groups_3D_LumVsOD_%s', timestamp);
        saveas(figCombined3D, fullfile(outputFolderPath, [combinedFilename '.png']));
        saveas(figCombined3D, fullfile(outputFolderPath, [combinedFilename '.fig']));
        
        % Update status
        set(handles.statusText, 'String', [
            '3D Lum vs. OD vs. Time analysis completed successfully!' newline newline ...
            'Results saved to: ' outputFolderPath newline ...
            'Generated files:' newline ...
            '• Individual group 3D trajectory plots' newline ...
            '• Combined 3D trajectory plot' newline ...
            '• Both .png and .fig formats saved' newline newline ...
            'Open .fig files in MATLAB to rotate and explore in 3D!' newline newline ...
            'Well selection window remains open for additional analyses.']);
        
    catch ME
        set(handles.statusText, 'String', ['Error during 3D analysis: ' ME.message newline newline ...
                                          'Please check your input files and try again.']);
        disp(getReport(ME, 'extended'));
    end
        end
        
        function launchEnhancedAnalysisGUI(groupData, groupAverage, groupStd, groupNames, timeVectorInHours, dataType, outputFolderPath, timestamp, textData, groupRows)
% Create enhanced analysis GUI window with FIXED positioning
    
    analysisWindow = figure('Name', sprintf('Enhanced %s Analysis - %s', dataType, timestamp), ...
                           'Position', [150, 150, 1200, 800], ...
                           'NumberTitle', 'off', 'MenuBar', 'none');
    
    % Create panels with FIXED positions
    controlPanel = uipanel(analysisWindow, 'Title', 'Analysis Options', ...
                          'Position', [0.02, 0.72, 0.28, 0.26], ...
                          'FontSize', 10, 'FontWeight', 'bold');
    
    % Plot panel (unchanged position)
    plotPanel = uipanel(analysisWindow, 'Title', 'Group Average Plot', ...
                       'Position', [0.32, 0.02, 0.66, 0.96], ...
                       'FontSize', 10, 'FontWeight', 'bold');
    
    % Parameters panel (unchanged position)
    parametersPanel = uipanel(analysisWindow, 'Title', 'Analysis Parameters', ...
                             'Position', [0.02, 0.02, 0.28, 0.68], ...
                             'FontSize', 10, 'FontWeight', 'bold');
    
    % Store data in the figure
    analysisData = struct();
    analysisData.groupData = groupData;
    analysisData.groupAverage = groupAverage;
    analysisData.groupStd = groupStd;
    analysisData.groupNames = groupNames;
    analysisData.timeVectorInHours = timeVectorInHours;
    analysisData.dataType = dataType;
    analysisData.outputFolderPath = outputFolderPath;
    analysisData.timestamp = timestamp;
    analysisData.textData = textData;
    analysisData.groupRows = groupRows;
    
    setappdata(analysisWindow, 'analysisData', analysisData);
    
    % Create analysis selection checkboxes with FIXED positioning (UNCHANGED)
    yStart = 180;  % Starting Y position
    spacing = 25;  % Spacing between checkboxes
    
    % Analysis options based on data type (UNCHANGED)
    if strcmp(dataType, 'OD')
        % OD-specific analyses
        analysisOptions = {
            'Growth Rate Analysis', 'growthRate';
            'Lag Time Analysis', 'lagTime';
            'Maximum OD Analysis', 'maxOD';
            'Doubling Time Analysis', 'doublingTime';
            'Area Under Curve', 'auc';
            'Growth Curve Fitting', 'curveFitting';
            'Statistical Comparisons', 'statistics';
            'Time to Thresholds', 'thresholds'
        };
    else
        % Luciferase-specific analyses
        analysisOptions = {
            'Peak Activity Analysis', 'peakActivity';
            'Expression Rate Analysis', 'expressionRate';
            'Total Expression (AUC)', 'auc';
            'Expression Duration', 'duration';
            'Fold Change Analysis', 'foldChange';
            'Expression Efficiency', 'efficiency';
            'Statistical Comparisons', 'statistics';
            'Temporal Patterns', 'patterns'
        };
    end
    
    % Create checkboxes for analysis options with FIXED positions (UNCHANGED)
    checkboxes = struct();
    for i = 1:length(analysisOptions)
        yPos = yStart - (i-1)*spacing;
        checkboxes.(analysisOptions{i, 2}) = uicontrol(controlPanel, 'Style', 'checkbox', ...
            'String', analysisOptions{i, 1}, ...
            'Position', [15, yPos, 260, 15], ...
            'Value', 1, 'FontSize', 9);
    end
    
    % Run Analysis button in control panel (UNCHANGED)
    runAnalysisBtn = uicontrol(controlPanel, 'Style', 'pushbutton', ...
                              'String', 'Run Selected Analyses', ...
                              'Position', [150, 15, 150, 35], ...
                              'FontWeight', 'bold', 'FontSize', 10, ...
                              'BackgroundColor', [0.2, 0.8, 0.2], ...
                              'ForegroundColor', 'white', ...
                              'Callback', @runEnhancedAnalysis);
    
    % Analysis parameters section with FIXED positioning (UNCHANGED)
paramYStart = 500;
paramSpacing = 35;

if strcmp(dataType, 'OD')
    % OD parameters with proper spacing
    % 1. OD Threshold for Lag Time
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'OD Threshold for Lag Time:', ...
             'Position', [15, paramYStart, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    lagThresholdEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '0.1', ...
                                'Position', [200, paramYStart, 60, 22], 'FontSize', 9);
    
    % 2. Growth Rate Start Time
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'Growth Rate Start Time (hrs):', ...
             'Position', [15, paramYStart-paramSpacing, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    growthStartEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '2', ...
                               'Position', [200, paramYStart-paramSpacing, 60, 22], 'FontSize', 9);
    
    % 3. Growth Rate End Time
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'Growth Rate End Time (hrs):', ...
             'Position', [15, paramYStart-2*paramSpacing, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    growthEndEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '6', ...
                             'Position', [200, paramYStart-2*paramSpacing, 60, 22], 'FontSize', 9);
    
    % 4. OD Thresholds (moved down to avoid overlap)
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'OD Thresholds (comma-sep):', ...
             'Position', [15, paramYStart-3*paramSpacing, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    thresholdsEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '0.1,0.5,1.0', ...
                              'Position', [15, paramYStart-3*paramSpacing-25, 245, 22], 'FontSize', 9);
    
    % Additional info text
    uicontrol(parametersPanel, 'Style', 'text', ...
             'String', 'Growth rate will be calculated between start and end times', ...
             'Position', [15, paramYStart-4*paramSpacing-25, 280, 40], ...
             'HorizontalAlignment', 'left', 'FontSize', 8, ...
             'ForegroundColor', [0.5, 0.5, 0.5]);
    
    paramControls = struct('lagThreshold', lagThresholdEdit, ...
                          'growthStart', growthStartEdit, ...
                          'growthEnd', growthEndEdit, ...
                          'thresholds', thresholdsEdit);
else
    % Luciferase parameters with proper spacing
    % 1. Baseline Threshold
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'Baseline Threshold (%):', ...
             'Position', [15, paramYStart, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    baselineEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '10', ...
                            'Position', [200, paramYStart, 60, 22], 'FontSize', 9);
    
    % 2. Control Group Index
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'Control Group Index:', ...
             'Position', [15, paramYStart-paramSpacing, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    controlGroupEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '1', ...
                                'Position', [200, paramYStart-paramSpacing, 60, 22], 'FontSize', 9);
    
    % 3. Expression Window
    uicontrol(parametersPanel, 'Style', 'text', 'String', 'Expression Window (hrs):', ...
             'Position', [15, paramYStart-2*paramSpacing, 180, 20], 'HorizontalAlignment', 'left', 'FontSize', 9);
    
    expressionWindowEdit = uicontrol(parametersPanel, 'Style', 'edit', 'String', '1', ...
                                    'Position', [200, paramYStart-2*paramSpacing, 60, 22], 'FontSize', 9);
    
    % Additional info text
    uicontrol(parametersPanel, 'Style', 'text', ...
             'String', 'Control group is used for fold change calculations', ...
             'Position', [15, paramYStart-3*paramSpacing-10, 280, 40], ...
             'HorizontalAlignment', 'left', 'FontSize', 8, ...
             'ForegroundColor', [0.5, 0.5, 0.5]);
    
    paramControls = struct('baseline', baselineEdit, ...
                          'controlGroup', controlGroupEdit, ...
                          'expressionWindow', expressionWindowEdit);
end
    
    setappdata(analysisWindow, 'checkboxes', checkboxes);
    setappdata(analysisWindow, 'paramControls', paramControls);
    
    % Plot axes for the average plot
    plotAxes = axes('Parent', plotPanel, 'Position', [0.08, 0.15, 0.85, 0.75]);
    
    % Create the initial average plot
    growth_analysis_module.createAveragePlot(plotAxes, groupAverage, groupStd, groupNames, timeVectorInHours, dataType);
    
    setappdata(analysisWindow, 'plotAxes', plotAxes);

    % Add a results display area
    resultsDisplay = uicontrol(plotPanel, 'Style', 'listbox', ...
                              'String', {'Analysis results will appear here...'}, ...
                              'Position', [0.05, 0.05, 0.9, 0.1], ... % Positioned at the bottom of plotPanel
                              'HorizontalAlignment', 'left', 'FontSize', 9, ...
                              'BackgroundColor', [0.98, 0.98, 0.98]);
    setappdata(analysisWindow, 'resultsDisplay', resultsDisplay); % Store for later use
    
    % Control buttons in plot panel - only Export and Generate Plots
    exportBtn = uicontrol(plotPanel, 'Style', 'pushbutton', ...
                         'String', 'Export Results', ...
                         'Position', [680, 15, 95, 30], ...
                         'FontSize', 9, 'Callback', @exportResults);
    
    plotBtn = uicontrol(plotPanel, 'Style', 'pushbutton', ...
                       'String', 'Generate Plots', ...
                       'Position', [575, 15, 95, 30], ...
                       'FontSize', 9, 'Callback', @generateAnalysisPlots);

    % Nested callback functions
    function runEnhancedAnalysis(src, ~)
        % Get stored data
        analysisData = getappdata(analysisWindow, 'analysisData');
        checkboxes = getappdata(analysisWindow, 'checkboxes');
        paramControls = getappdata(analysisWindow, 'paramControls');
        resultsDisplay = getappdata(analysisWindow, 'resultsDisplay'); % Get results display handle
        
        % Clear previous results
        set(resultsDisplay, 'String', {'Running analysis...'});
        drawnow;

        % Store results for export
        results = struct();
        results.summary = {};
        results.tables = {};
        results.statistics = {};
        
        try
            if strcmp(analysisData.dataType, 'OD')
                results = growth_analysis_module.runODAnalyses(analysisData, checkboxes, paramControls, resultsDisplay, results);
            else
                results = growth_analysis_module.runLuciferaseAnalyses(analysisData, checkboxes, paramControls, resultsDisplay, results);
            end
            
            % Store results for export
            setappdata(analysisWindow, 'analysisResults', results);
            
            % Update status in results display
try
    growth_analysis_module.addResult(resultsDisplay, ' ');
    growth_analysis_module.addResult(resultsDisplay, 'Analysis completed successfully! Use Export/Generate Plots buttons for full results.');
catch
    % Fallback if resultsDisplay handle is broken
    set(resultsDisplay, 'String', [get(resultsDisplay, 'String'); {' '}; {'Analysis completed successfully! Use Export/Generate Plots buttons for full results.'}]);
end

        catch ME
            growth_analysis_module.addResult(resultsDisplay, sprintf('Error during analysis: %s', ME.message));
            disp(getReport(ME, 'extended'));
        end
    end

    function exportResults(src, ~)
        growth_analysis_module.exportResults(src, []);
    end

    function generateAnalysisPlots(src, ~)
        growth_analysis_module.generateAnalysisPlots(src, []);
    end
        end
        
        function createAveragePlot(ax, groupAverage, groupStd, groupNames, timeVectorInHours, dataType)
% Clear axes
    cla(ax);
    
    % Generate colors for groups
    numGroups = length(groupNames);
    groupColors = zeros(numGroups, 3);
    for i = 1:numGroups
        groupColors(i, :) = growth_analysis_module.generateUniqueColor(i);
    end
    
    % Plot data
    hold(ax, 'on');
    
    % Plot standard deviations as filled regions - FIXED: robust fill function
    for i = 1:numGroups
        if ~isempty(groupAverage{i})
            try
                % FIXED: Ensure proper dimensions and no negative values
                avgData = groupAverage{i};
                stdData = groupStd{i};
                timeData = timeVectorInHours;
                
                % Ensure all are row vectors
                if size(avgData, 1) > size(avgData, 2)
                    avgData = avgData';
                end
                if size(stdData, 1) > size(stdData, 2)
                    stdData = stdData';
                end
                if size(timeData, 1) > size(timeData, 2)
                    timeData = timeData';
                end
                
                % Remove any NaN or infinite values
                validIdx = isfinite(avgData) & isfinite(stdData) & isfinite(timeData);
                if any(validIdx)
                    avgData = avgData(validIdx);
                    stdData = stdData(validIdx);
                    timeData = timeData(validIdx);
                    
                    % Calculate bounds
                    lowerBound = max(0, avgData - stdData);
                    upperBound = avgData + stdData;
                    
                    % Create fill vectors
                    fillX = [timeData, fliplr(timeData)];
                    fillY = [lowerBound, fliplr(upperBound)];
                    
                    % Only fill if we have valid data
                    if length(fillX) > 2 && length(fillY) > 2
                        % Try fill first, fallback to patch if it fails
                        try
                            fill(ax, fillX, fillY, groupColors(i, :), 'FaceAlpha', 0.3, 'EdgeColor', 'none');
                        catch
                            % Fallback to patch
                            patch(ax, fillX, fillY, groupColors(i, :), 'FaceAlpha', 0.3, 'EdgeColor', 'none');
                        end
                    end
                end
            catch ME
                fprintf('Fill error for group %d in createAveragePlot: %s\n', i, ME.message);
                % Continue without fill for this group
            end
        end
    end
    
    % Plot the averages
    h = gobjects(1, numGroups);
    for i = 1:numGroups
        if ~isempty(groupAverage{i})
            h(i) = plot(ax, timeVectorInHours, groupAverage{i}, ...
                        'LineWidth', 2, 'Color', groupColors(i, :));
        end
    end
    
    % Formatting
    if strcmp(dataType, 'OD')
        title(ax, 'Average OD Growth Curves', 'FontSize', 14, 'FontWeight', 'bold');
        ylabel(ax, 'OD 600nm', 'FontSize', 12, 'FontWeight', 'bold');
    else
        title(ax, 'Average Normalized Luciferase Expression', 'FontSize', 14, 'FontWeight', 'bold');
        ylabel(ax, 'Normalized Luciferase', 'FontSize', 12, 'FontWeight', 'bold');
    end
    
    xlabel(ax, 'Time (hours)', 'FontSize', 12, 'FontWeight', 'bold');
legend(ax, h, groupNames, 'Location', 'best', 'FontSize', 10);
grid(ax, 'on');
set(ax, 'FontSize', 10, 'LineWidth', 1.5);

% Set x-axis to show half-hour intervals
xlimits = get(ax, 'XLim');
xticks = 0:1:ceil(xlimits(2)); % hour intervals from 0 to max time
set(ax, 'XTick', xticks);
set(ax, 'XTickLabelRotation', 45); % Rotate labels if they overlap

hold(ax, 'off');
        end

        
        function results = runODAnalyses(analysisData, checkboxes, paramControls, resultsText, results)
% Run OD-specific analyses - FIXED VERSION
    
    groupData = analysisData.groupData;
    groupAverage = analysisData.groupAverage;
    groupNames = analysisData.groupNames;
    timeVectorInHours = analysisData.timeVectorInHours;
    
    numGroups = length(groupNames);
    
    % Get parameters
    lagThreshold = str2double(get(paramControls.lagThreshold, 'String'));
    growthStart = str2double(get(paramControls.growthStart, 'String'));
    growthEnd = str2double(get(paramControls.growthEnd, 'String'));
    thresholds = str2num(get(paramControls.thresholds, 'String')); %#ok<ST2NM>
    
    % Validate parameters
    if ~isfinite(lagThreshold) || lagThreshold <= 0
        lagThreshold = 0.1; % Default threshold
    end
    
    if ~isfinite(growthStart) || growthStart < 0
        growthStart = 2; % Default start time
    end
    
    if ~isfinite(growthEnd) || growthEnd <= growthStart
        growthEnd = growthStart + 4; % Default 4 hours after start
    end
    
    if isempty(thresholds) || any(~isfinite(thresholds)) || any(thresholds <= 0)
        thresholds = [0.1, 0.5, 1.0]; % Default thresholds
    end
    
    % Initialize variables for dependencies
    growthRates = [];
    growthRateStd = [];
    aucValues = [];
    
    % Helper function for safe result logging
    addResultSafe = @(text) [];
    if ~isempty(resultsText)
        addResultSafe = @(text) growth_analysis_module.addResult(resultsText, text);
    end
    
    % 1. Growth Rate Analysis
    if get(checkboxes.growthRate, 'Value')
        addResultSafe('--- GROWTH RATE ANALYSIS ---');
        
        growthRates = zeros(numGroups, 1);
        growthRateStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
    if ~isempty(groupData{i})
        rates = [];
        for j = 1:size(groupData{i}, 1)
            rate = growth_analysis_module.calculateGrowthRate(timeVectorInHours, groupData{i}(j, :), growthStart, growthEnd);
            rates = [rates, rate];
        end
        growthRates(i) = mean(rates);
        growthRateStd(i) = std(rates);
        
        addResultSafe(sprintf('%s: %.4f ± %.4f h⁻¹', ...
                 groupNames{i}, growthRates(i), growthRateStd(i)));
    end
        end
    
        
        results.growthRates = table(groupNames', growthRates, growthRateStd, ...
                                   'VariableNames', {'Group', 'GrowthRate_per_h', 'StdDev'});
        addResultSafe('');
    end
    
    % 2. Lag Time Analysis
    if get(checkboxes.lagTime, 'Value')
        addResultSafe('--- LAG TIME ANALYSIS ---');
        
        lagTimes = zeros(numGroups, 1);
        lagTimeStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                lags = [];
                for j = 1:size(groupData{i}, 1)
                    lag = growth_analysis_module.calculateLagTime(timeVectorInHours, groupData{i}(j, :), lagThreshold);
                    lags = [lags, lag];
                end
                lagTimes(i) = mean(lags);
                lagTimeStd(i) = std(lags);
                
                addResultSafe(sprintf('%s: %.2f ± %.2f hours', ...
                         groupNames{i}, lagTimes(i), lagTimeStd(i)));
            end
        end
        
        results.lagTimes = table(groupNames', lagTimes, lagTimeStd, ...
                                'VariableNames', {'Group', 'LagTime_h', 'StdDev'});
        addResultSafe('');
    end
    
    % 3. Maximum OD Analysis
    if get(checkboxes.maxOD, 'Value')
        addResultSafe('--- MAXIMUM OD ANALYSIS ---');
        
        maxODs = zeros(numGroups, 1);
        maxODStd = zeros(numGroups, 1);
        timeToMax = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                maxVals = max(groupData{i}, [], 2);
                maxODs(i) = mean(maxVals);
                maxODStd(i) = std(maxVals);
                
                [~, maxIdx] = max(groupAverage{i});
                timeToMax(i) = timeVectorInHours(maxIdx);
                
                addResultSafe(sprintf('%s: %.3f ± %.3f (at %.1f h)', ...
                         groupNames{i}, maxODs(i), maxODStd(i), timeToMax(i)));
            end
        end
        
        results.maxOD = table(groupNames', maxODs, maxODStd, timeToMax, ...
                             'VariableNames', {'Group', 'MaxOD', 'StdDev', 'TimeToMax_h'});
        addResultSafe('');
    end
    
    % 4. Doubling Time Analysis (depends on growth rates)
    if get(checkboxes.doublingTime, 'Value')
        addResultSafe('--- DOUBLING TIME ANALYSIS ---');
        
        % Calculate growth rates if not already done
        if isempty(growthRates)
            growthRates = zeros(numGroups, 1);
            growthRateStd = zeros(numGroups, 1);
            
            for i = 1:numGroups
    if ~isempty(groupData{i})
        rates = [];
        for j = 1:size(groupData{i}, 1)
            rate = growth_analysis_module.calculateGrowthRate(timeVectorInHours, groupData{i}(j, :), growthStart, growthEnd);
            rates = [rates, rate];
        end
        growthRates(i) = mean(rates);
        growthRateStd(i) = std(rates);
    end
            end
        end
        
        doublingTimes = zeros(numGroups, 1);
        doublingTimeStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i}) && growthRates(i) > 0
                dt = log(2) / growthRates(i);
                doublingTimes(i) = dt;
                
                % Calculate std based on growth rate std
                if growthRateStd(i) > 0
                    doublingTimeStd(i) = log(2) * growthRateStd(i) / (growthRates(i)^2);
                end
                
                addResultSafe(sprintf('%s: %.2f ± %.2f hours', ...
                         groupNames{i}, doublingTimes(i), doublingTimeStd(i)));
            end
        end
        
        results.doublingTimes = table(groupNames', doublingTimes, doublingTimeStd, ...
                                     'VariableNames', {'Group', 'DoublingTime_h', 'StdDev'});
        addResultSafe('');
    end
    
    % 5. Area Under Curve Analysis
    if get(checkboxes.auc, 'Value')
        addResultSafe('--- AREA UNDER CURVE ANALYSIS ---');
        
        aucValues = zeros(numGroups, 1);
        aucStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                aucs = [];
                for j = 1:size(groupData{i}, 1)
                    auc = trapz(timeVectorInHours, groupData{i}(j, :));
                    aucs = [aucs, auc];
                end
                aucValues(i) = mean(aucs);
                aucStd(i) = std(aucs);
                
                addResultSafe(sprintf('%s: %.2f ± %.2f OD·h', ...
                         groupNames{i}, aucValues(i), aucStd(i)));
            end
        end
        
        results.auc = table(groupNames', aucValues, aucStd, ...
                           'VariableNames', {'Group', 'AUC_OD_h', 'StdDev'});
        addResultSafe('');
    end
    
    % 6. Growth Curve Fitting
    if get(checkboxes.curveFitting, 'Value')
        addResultSafe('--- GROWTH CURVE FITTING ---');
        
        rSquared = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                % Fit logistic growth model to average curve
                try
                    % Simple exponential fit for demonstration
                    validIdx = groupAverage{i} > 0;
                    if sum(validIdx) > 3
                        logData = log(groupAverage{i}(validIdx));
                        timeValid = timeVectorInHours(validIdx);
                        p = polyfit(timeValid, logData, 1);
                        
                        % Calculate R-squared
                        yfit = polyval(p, timeValid);
                        ssr = sum((logData - yfit).^2);
                        sst = sum((logData - mean(logData)).^2);
                        rSquared(i) = 1 - ssr/sst;
                        
                        addResultSafe(sprintf('%s: R² = %.3f', ...
                                 groupNames{i}, rSquared(i)));
                    end
                catch
                    rSquared(i) = NaN;
                    addResultSafe(sprintf('%s: Fitting failed', groupNames{i}));
                end
            end
        end
        
        results.curveFitting = table(groupNames', rSquared, ...
                                    'VariableNames', {'Group', 'R_squared'});
        addResultSafe('');
    end
    
    % 7. Time to Thresholds Analysis
    if get(checkboxes.thresholds, 'Value')
        addResultSafe('--- TIME TO THRESHOLDS ANALYSIS ---');
        
        thresholdResults = {};
        for t = 1:length(thresholds)
            thresh = thresholds(t);
            times = zeros(numGroups, 1);
            timeStd = zeros(numGroups, 1);
            
            addResultSafe(sprintf('Threshold OD = %.1f:', thresh));
            
            for i = 1:numGroups
                if ~isempty(groupData{i})
                    threshTimes = [];
                    for j = 1:size(groupData{i}, 1)
                        idx = find(groupData{i}(j, :) >= thresh, 1);
                        if ~isempty(idx)
                            threshTimes = [threshTimes, timeVectorInHours(idx)];
                        end
                    end
                    if ~isempty(threshTimes)
                        times(i) = mean(threshTimes);
                        timeStd(i) = std(threshTimes);
                        addResultSafe(sprintf('  %s: %.2f ± %.2f hours', ...
                                 groupNames{i}, times(i), timeStd(i)));
                    else
                        addResultSafe(sprintf('  %s: Threshold not reached', groupNames{i}));
                    end
                end
            end
            
            thresholdResults{t} = table(groupNames', times, timeStd, ...
                                       'VariableNames', {'Group', sprintf('Time_to_OD%.1f_h', thresh), 'StdDev'});
        end
        results.thresholds = thresholdResults;
        addResultSafe('');
    end
    
    % 8. Statistical Comparisons
    if get(checkboxes.statistics, 'Value')
        addResultSafe('--- STATISTICAL COMPARISONS ---');
        
        if numGroups >= 2
            % Calculate AUC if not already done
            if isempty(aucValues)
                aucValues = zeros(numGroups, 1);
                for i = 1:numGroups
                    if ~isempty(groupData{i})
                        aucs = [];
                        for j = 1:size(groupData{i}, 1)
                            auc = trapz(timeVectorInHours, groupData{i}(j, :));
                            aucs = [aucs, auc];
                        end
                        aucValues(i) = mean(aucs);
                    end
                end
            end
            
            % Perform ANOVA on AUC values
            [p_auc, tbl_auc] = growth_analysis_module.performANOVA(groupData, @(data) trapz(timeVectorInHours, data));
            addResultSafe(sprintf('ANOVA (AUC): F = %.3f, p = %.4f', tbl_auc{2,5}, p_auc));
            
            % Pairwise comparisons
            if p_auc < 0.05
                addResultSafe('Significant differences detected. Pairwise comparisons:');
                for i = 1:numGroups-1
                    for j = i+1:numGroups
                        if ~isempty(groupData{i}) && ~isempty(groupData{j})
                            aucs_i = [];
                            aucs_j = [];
                            for k = 1:size(groupData{i}, 1)
                                aucs_i = [aucs_i, trapz(timeVectorInHours, groupData{i}(k, :))];
                            end
                            for k = 1:size(groupData{j}, 1)
                                aucs_j = [aucs_j, trapz(timeVectorInHours, groupData{j}(k, :))];
                            end
                            [~, p_pair] = ttest2(aucs_i, aucs_j);
                            addResultSafe(sprintf('  %s vs %s: p = %.4f', ...
                                     groupNames{i}, groupNames{j}, p_pair));
                        end
                    end
                end
            end
        else
            addResultSafe('Need at least 2 groups for statistical comparisons');
        end
        addResultSafe('');
    end
        end
        
        function results = runLuciferaseAnalyses(analysisData, checkboxes, paramControls, resultsText, results)
% Run Luciferase-specific analyses
    
    groupData = analysisData.groupData;
    groupAverage = analysisData.groupAverage;
    groupNames = analysisData.groupNames;
    timeVectorInHours = analysisData.timeVectorInHours;
    
    numGroups = length(groupNames);
    
    % Get parameters
    baseline = str2double(get(paramControls.baseline, 'String')) / 100;
    controlGroupIdx = str2double(get(paramControls.controlGroup, 'String'));
    expressionWindow = str2double(get(paramControls.expressionWindow, 'String'));
    
    % Validate parameters
    if ~isfinite(baseline) || baseline < 0 || baseline > 1
        baseline = 0.1; % Default to 10%
    end
    
    if ~isfinite(controlGroupIdx) || controlGroupIdx < 1 || controlGroupIdx > numGroups
        controlGroupIdx = 1; % Default to first group
    end
    
    if ~isfinite(expressionWindow) || expressionWindow <= 0
        expressionWindow = 1; % Default to 1 hour
    end
    
    % Helper function for safe result logging
    addResultSafe = @(text) [];
    if ~isempty(resultsText)
        addResultSafe = @(text) growth_analysis_module.addResult(resultsText, text);
    end
    
    % Peak Activity Analysis
    if get(checkboxes.peakActivity, 'Value')
        addResultSafe('--- PEAK ACTIVITY ANALYSIS ---');
        
        peakValues = zeros(numGroups, 1);
        peakStd = zeros(numGroups, 1);
        timeToPeak = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                peaks = max(groupData{i}, [], 2);
                peakValues(i) = mean(peaks);
                peakStd(i) = std(peaks);
                
                [~, maxIdx] = max(groupAverage{i});
                timeToPeak(i) = timeVectorInHours(maxIdx);
                
                addResultSafe(sprintf('%s: %.2f ± %.2f (at %.1f h)', ...
                         groupNames{i}, peakValues(i), peakStd(i), timeToPeak(i)));
            end
        end
        
        results.peakActivity = table(groupNames', peakValues, peakStd, timeToPeak, ...
                                    'VariableNames', {'Group', 'PeakActivity', 'StdDev', 'TimeToPeak_h'});
        addResultSafe('');
    end
    
    % Expression Rate Analysis
    if get(checkboxes.expressionRate, 'Value')
        addResultSafe('--- EXPRESSION RATE ANALYSIS ---');
        
        expressionRates = zeros(numGroups, 1);
        rateStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                rates = [];
                for j = 1:size(groupData{i}, 1)
                    rate = growth_analysis_module.calculateExpressionRate(timeVectorInHours, groupData{i}(j, :), expressionWindow);
                    rates = [rates, rate];
                end
                expressionRates(i) = mean(rates);
                rateStd(i) = std(rates);
                
                addResultSafe(sprintf('%s: %.3f ± %.3f units/h', ...
                         groupNames{i}, expressionRates(i), rateStd(i)));
            end
        end
        
        results.expressionRates = table(groupNames', expressionRates, rateStd, ...
                                       'VariableNames', {'Group', 'ExpressionRate_units_per_h', 'StdDev'});
        addResultSafe('');
    end
    
    % Total Expression (AUC) Analysis
    if get(checkboxes.auc, 'Value')
        addResultSafe('--- TOTAL EXPRESSION (AUC) ANALYSIS ---');
        
        aucValues = zeros(numGroups, 1);
        aucStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                aucs = [];
                for j = 1:size(groupData{i}, 1)
                    auc = trapz(timeVectorInHours, groupData{i}(j, :));
                    aucs = [aucs, auc];
                end
                aucValues(i) = mean(aucs);
                aucStd(i) = std(aucs);
                
                addResultSafe(sprintf('%s: %.2f ± %.2f units·h', ...
                         groupNames{i}, aucValues(i), aucStd(i)));
            end
        end
        
        results.auc = table(groupNames', aucValues, aucStd, ...
                           'VariableNames', {'Group', 'AUC_units_h', 'StdDev'});
        addResultSafe('');
    end
    
    % Expression Duration Analysis
    if get(checkboxes.duration, 'Value')
        addResultSafe('--- EXPRESSION DURATION ANALYSIS ---');
        
        durations = zeros(numGroups, 1);
        durationStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                durs = [];
                for j = 1:size(groupData{i}, 1)
                    dur = growth_analysis_module.calculateExpressionDuration(timeVectorInHours, groupData{i}(j, :), baseline);
                    durs = [durs, dur];
                end
                durations(i) = mean(durs);
                durationStd(i) = std(durs);
                
                addResultSafe(sprintf('%s: %.2f ± %.2f hours', ...
                         groupNames{i}, durations(i), durationStd(i)));
            end
        end
        
        results.duration = table(groupNames', durations, durationStd, ...
                                'VariableNames', {'Group', 'Duration_h', 'StdDev'});
        addResultSafe('');
    end
    
    % Fold Change Analysis
    if get(checkboxes.foldChange, 'Value')
        addResultSafe('--- FOLD CHANGE ANALYSIS ---');
        
        if controlGroupIdx > 0 && controlGroupIdx <= numGroups && ~isempty(groupData{controlGroupIdx})
            controlPeak = mean(max(groupData{controlGroupIdx}, [], 2));
            
            foldChanges = zeros(numGroups, 1);
            fcStd = zeros(numGroups, 1);
            
            for i = 1:numGroups
                if ~isempty(groupData{i})
                    peaks = max(groupData{i}, [], 2);
                    fcs = peaks / controlPeak;
                    foldChanges(i) = mean(fcs);
                    fcStd(i) = std(fcs);
                    
                    addResultSafe(sprintf('%s: %.2f ± %.2f fold', ...
                             groupNames{i}, foldChanges(i), fcStd(i)));
                end
            end
            
            results.foldChange = table(groupNames', foldChanges, fcStd, ...
                                      'VariableNames', {'Group', 'FoldChange', 'StdDev'});
        else
            addResultSafe('Invalid control group specified');
        end
        addResultSafe('');
    end
    
    % Expression Efficiency Analysis
    if get(checkboxes.efficiency, 'Value')
        addResultSafe('--- EXPRESSION EFFICIENCY ANALYSIS ---');
        
        efficiencies = zeros(numGroups, 1);
        effStd = zeros(numGroups, 1);
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                effs = [];
                for j = 1:size(groupData{i}, 1)
                    % Efficiency = Peak / Time to Peak
                    [peak, peakIdx] = max(groupData{i}(j, :));
                    timeToPeak = timeVectorInHours(peakIdx);
                    if timeToPeak > 0
                        eff = peak / timeToPeak;
                        effs = [effs, eff];
                    end
                end
                if ~isempty(effs)
                    efficiencies(i) = mean(effs);
                    effStd(i) = std(effs);
                    
                    addResultSafe(sprintf('%s: %.3f ± %.3f units/h', ...
                             groupNames{i}, efficiencies(i), effStd(i)));
                end
            end
        end
        
        results.efficiency = table(groupNames', efficiencies, effStd, ...
                                  'VariableNames', {'Group', 'Efficiency_units_per_h', 'StdDev'});
        addResultSafe('');
    end
    
    % Statistical Comparisons
    if get(checkboxes.statistics, 'Value')
        addResultSafe('--- STATISTICAL COMPARISONS ---');
        
        if numGroups >= 2
            % Perform ANOVA on peak values
            if exist('peakValues', 'var')
                [p_peak, tbl_peak] = growth_analysis_module.performANOVA(groupData, @(data) max(data));
                addResultSafe(sprintf('ANOVA (Peak): F = %.3f, p = %.4f', tbl_peak{2,5}, p_peak));
                
                % Pairwise comparisons
                if p_peak < 0.05
                    addResultSafe('Significant differences detected. Pairwise comparisons:');
                    for i = 1:numGroups-1
                        for j = i+1:numGroups
                            if ~isempty(groupData{i}) && ~isempty(groupData{j})
                                peaks_i = max(groupData{i}, [], 2);
                                peaks_j = max(groupData{j}, [], 2);
                                [~, p_pair] = ttest2(peaks_i, peaks_j);
                                addResultSafe(sprintf('  %s vs %s: p = %.4f', ...
                                         groupNames{i}, groupNames{j}, p_pair));
                            end
                        end
                    end
                end
            end
        else
            addResultSafe('Need at least 2 groups for statistical comparisons');
        end
        addResultSafe('');
    end
    
    % Temporal Patterns Analysis
    if get(checkboxes.patterns, 'Value')
        addResultSafe('--- TEMPORAL EXPRESSION PATTERNS ---');
        
        for i = 1:numGroups
            if ~isempty(groupData{i})
                avg_curve = groupAverage{i};
                
                % Find onset time (first significant increase)
                baseline_val = mean(avg_curve(1:min(3, end)));
                onset_idx = find(avg_curve > baseline_val * (1 + baseline), 1);
                onset_time = onset_idx * mean(diff(timeVectorInHours));
                
                % Find offset time (return to baseline)
                [peak_val, peak_idx] = max(avg_curve);
                post_peak = avg_curve(peak_idx:end);
                offset_idx = find(post_peak < peak_val * baseline, 1);
                if ~isempty(offset_idx)
                    offset_time = (peak_idx + offset_idx - 1) * mean(diff(timeVectorInHours));
                else
                    offset_time = timeVectorInHours(end);
                end
                
                addResultSafe(sprintf('%s:', groupNames{i}));
                addResultSafe(sprintf('  Onset: %.1f h, Peak: %.1f h, Offset: %.1f h', ...
                         onset_time, timeVectorInHours(peak_idx), offset_time));
                addResultSafe(sprintf('  Active duration: %.1f h', offset_time - onset_time));
            end
        end
        addResultSafe('');
    end
        end
        
        function rate = calculateGrowthRate(time, data, startTime, endTime)
% Calculate growth rate in specified time range
    
    % Input validation
    if isempty(time) || isempty(data) || length(time) ~= length(data)
        rate = NaN;
        return;
    end
    
    if ~isfinite(startTime) || ~isfinite(endTime) || startTime >= endTime
        rate = NaN;
        return;
    end
    
    % Find indices for the time range
    startIdx = find(time >= startTime, 1, 'first');
    endIdx = find(time <= endTime, 1, 'last');
    
    if isempty(startIdx) || isempty(endIdx) || startIdx >= endIdx
        rate = NaN;
        return;
    end
    
    % Extract data in the time range
    timeRange = time(startIdx:endIdx);
    dataRange = data(startIdx:endIdx);
    
    % Remove any NaN/Inf values
    validIdx = isfinite(timeRange) & isfinite(dataRange) & dataRange > 0;
    if sum(validIdx) < 2
        rate = NaN;
        return;
    end
    
    timeValid = timeRange(validIdx);
    dataValid = dataRange(validIdx);
    
    try
        % Log transform for exponential growth
        logData = log(dataValid);
        
        % Check for valid log data
        if any(~isfinite(logData))
            rate = NaN;
            return;
        end
        
        % Linear fit to get growth rate
        p = polyfit(timeValid, logData, 1);
        
        if isfinite(p(1))
            rate = p(1); % Slope is the growth rate
        else
            rate = NaN;
        end
    catch
        rate = NaN;
    end
   end
        
        function lag = calculateLagTime(time, data, threshold)
% Calculate lag time as time to reach threshold
    
    % Input validation
    if isempty(time) || isempty(data) || length(time) ~= length(data)
        lag = NaN;
        return;
    end
    
    if ~isfinite(threshold) || threshold <= 0
        lag = NaN;
        return;
    end
    
    % Remove NaN/Inf values
    validIdx = isfinite(time) & isfinite(data);
    if ~any(validIdx)
        lag = NaN;
        return;
    end
    
    time = time(validIdx);
    data = data(validIdx);
    
    % Find first point above threshold
    idx = find(data >= threshold, 1);
    if ~isempty(idx) && idx <= length(time)
        lag = time(idx);
    else
        lag = NaN;  % Threshold never reached
    end
        end
        
        function rate = calculateExpressionRate(time, data, window)
% Calculate maximum expression rate using a sliding window
    
    % Input validation
    if isempty(time) || isempty(data) || length(time) ~= length(data)
        rate = NaN;
        return;
    end
    
    if ~isfinite(window) || window <= 0
        rate = NaN;
        return;
    end
    
    % Remove NaN/Inf values
    validIdx = isfinite(time) & isfinite(data);
    if sum(validIdx) < 3  % Need at least 3 points
        rate = NaN;
        return;
    end
    
    time = time(validIdx);
    data = data(validIdx);
    
    % Calculate time step
    timeDiff = diff(time);
    if isempty(timeDiff) || any(~isfinite(timeDiff))
        rate = NaN;
        return;
    end
    
    meanTimeStep = mean(timeDiff);
    if ~isfinite(meanTimeStep) || meanTimeStep <= 0
        rate = NaN;
        return;
    end
    
    % Calculate window points
    windowPoints = round(window / meanTimeStep);
    
    % Validate window points
    if ~isfinite(windowPoints) || windowPoints < 1
        windowPoints = max(1, min(3, length(data) - 1));  % Fallback to small window
    end
    
    if windowPoints >= length(data)
        windowPoints = max(1, length(data) - 1);
    end
    
    % Calculate rates
    rates = [];
    maxIdx = length(data) - windowPoints;
    
    if maxIdx < 1
        % Too few points, just calculate overall rate
        if length(time) >= 2
            p = polyfit(time, data, 1);
            rate = p(1);
        else
            rate = NaN;
        end
        return;
    end
    
    for i = 1:maxIdx
        try
            endIdx = min(i + windowPoints, length(data));
            subset_time = time(i:endIdx);
            subset_data = data(i:endIdx);
            
            if length(subset_time) >= 2 && all(isfinite(subset_time)) && all(isfinite(subset_data))
                % Linear fit
                p = polyfit(subset_time, subset_data, 1);
                if isfinite(p(1))
                    rates = [rates, p(1)]; % Slope is expression rate
                end
            end
        catch
            % Skip this window if there's an error
            continue;
        end
    end
    
    if isempty(rates)
        rate = NaN;
    else
        rate = max(rates); % Maximum expression rate
    end
     end
        
        function duration = calculateExpressionDuration(time, data, baselineThreshold)
 % Calculate expression duration above baseline
    
    % Input validation
    if isempty(time) || isempty(data) || length(time) ~= length(data)
        duration = 0;
        return;
    end
    
    if ~isfinite(baselineThreshold) || baselineThreshold < 0 || baselineThreshold > 1
        duration = 0;
        return;
    end
    
    % Remove NaN/Inf values
    validIdx = isfinite(time) & isfinite(data);
    if sum(validIdx) < 2
        duration = 0;
        return;
    end
    
    time = time(validIdx);
    data = data(validIdx);
    
    % Calculate threshold
    maxVal = max(data);
    if ~isfinite(maxVal) || maxVal <= 0
        duration = 0;
        return;
    end
    
    threshold = maxVal * baselineThreshold;
    
    % Find points above threshold
    aboveThreshold = data > threshold;
    if any(aboveThreshold)
        aboveIndices = find(aboveThreshold);
        firstIdx = aboveIndices(1);
        lastIdx = aboveIndices(end);
        
        if firstIdx <= length(time) && lastIdx <= length(time) && lastIdx > firstIdx
            duration = time(lastIdx) - time(firstIdx);
        else
            duration = 0;
        end
    else
        duration = 0;
    end
        end
        
        function [p, tbl] = performANOVA(groupData, metricFunc)
% Perform ANOVA on a specific metric
    
    allData = [];
    groupLabels = [];
    
    try
        for i = 1:length(groupData)
            if ~isempty(groupData{i})
                for j = 1:size(groupData{i}, 1)
                    metric = metricFunc(groupData{i}(j, :));
                    if isfinite(metric)  % Only include finite values
                        allData = [allData; metric];
                        groupLabels = [groupLabels; i];
                    end
                end
            end
        end
        
        % Check if we have enough data for ANOVA
        if length(allData) < 3 || length(unique(groupLabels)) < 2
            p = NaN;
            tbl = {};
            return;
        end
        
        [p, tbl] = anova1(allData, groupLabels, 'off');
        
        % Validate results
        if ~isfinite(p)
            p = NaN;
        end
        
    catch ME
        fprintf('ANOVA failed: %s\n', ME.message);
        p = NaN;
        tbl = {};
    end
        end
        
        function addResult(resultsText, newLine)
% Add a new line to the results display
    
    currentLines = get(resultsText, 'String');
    if ischar(currentLines)
        currentLines = {currentLines};
    end
    
    newLines = [currentLines; {newLine}];
    set(resultsText, 'String', newLines);
    set(resultsText, 'Value', length(newLines)); % Scroll to bottom
        end

        function exportResults(src, ~)
% Export analysis results to Excel
    
    analysisWindow = get(src, 'Parent');
    while ~strcmp(get(analysisWindow, 'Type'), 'figure')
        analysisWindow = get(analysisWindow, 'Parent');
    end
    
    analysisData = getappdata(analysisWindow, 'analysisData');
    analysisResults = getappdata(analysisWindow, 'analysisResults');
    
    if isempty(analysisResults)
        msgbox('No results to export. Please run analysis first.', 'No Results');
        return;
    end
    
    % Create export filename
    exportFilename = fullfile(analysisData.outputFolderPath, ...
                             sprintf('Enhanced_%s_Analysis_%s.xlsx', ...
                             analysisData.dataType, analysisData.timestamp));
    
    try
        % Write each table to a separate sheet
        fieldNames = fieldnames(analysisResults);
        
        for i = 1:length(fieldNames)
            fieldName = fieldNames{i};
            if istable(analysisResults.(fieldName))
                writetable(analysisResults.(fieldName), exportFilename, ...
                          'Sheet', fieldName, 'WriteMode', 'overwrite');
            elseif iscell(analysisResults.(fieldName))
                % Handle threshold results (cell array of tables)
                for j = 1:length(analysisResults.(fieldName))
                    if istable(analysisResults.(fieldName){j})
                        sheetName = sprintf('%s_%d', fieldName, j);
                        writetable(analysisResults.(fieldName){j}, exportFilename, ...
                                  'Sheet', sheetName, 'WriteMode', 'overwrite');
                    end
                end
            end
        end
        
        msgbox(sprintf('Results exported to:\n%s', exportFilename), 'Export Complete');
        
    catch ME
        msgbox(sprintf('Error exporting results:\n%s', ME.message), 'Export Error');
    end
        end
        
         function generateAnalysisPlots(src, ~)
% Generate ALL enhanced analysis plots including the 6 working bar charts
    
    analysisWindow = get(src, 'Parent');
    while ~strcmp(get(analysisWindow, 'Type'), 'figure')
        analysisWindow = get(analysisWindow, 'Parent');
    end
    
    analysisData = getappdata(analysisWindow, 'analysisData');
    analysisResults = getappdata(analysisWindow, 'analysisResults');
    
    if isempty(analysisResults)
        msgbox('No results to plot. Please run analysis first.', 'No Results');
        return;
    end
    
    try
        timestamp = analysisData.timestamp;
        outputPath = analysisData.outputFolderPath;
        dataType = analysisData.dataType;
        groupNames = analysisData.groupNames;
        numGroups = length(groupNames);
        
        % Generate group colors using the same function
        groupColors = zeros(numGroups, 3);
        for i = 1:numGroups
            groupColors(i, :) = growth_analysis_module.generateUniqueColor(i);
        end
        
        % PLOT 1: Main Summary Bar Charts (the 6 working ones)
        if isfield(analysisResults, 'growthRates') || isfield(analysisResults, 'peakActivity')
            fig1 = figure('Name', sprintf('%s Summary Bar Charts - %s', dataType, timestamp), ...
                         'Position', [100, 100, 1200, 800]);
            
            if strcmp(dataType, 'OD')
                % OD Analysis Bar Charts
                plotCount = 1;
                
                % 1. Growth Rate
                if isfield(analysisResults, 'growthRates')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.growthRates.GrowthRate_per_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Growth Rate (h⁻¹)', 'FontWeight', 'bold');
                    ylabel('Growth Rate (h⁻¹)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 2. Maximum OD
                if isfield(analysisResults, 'maxOD')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.maxOD.MaxOD);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Maximum OD', 'FontWeight', 'bold');
                    ylabel('Max OD');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 3. Lag Time
                if isfield(analysisResults, 'lagTimes')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.lagTimes.LagTime_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Lag Time (h)', 'FontWeight', 'bold');
                    ylabel('Lag Time (h)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 4. AUC
                if isfield(analysisResults, 'auc')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.auc.AUC_OD_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Area Under Curve', 'FontWeight', 'bold');
                    ylabel('AUC (OD·h)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 5. Doubling Time
                if isfield(analysisResults, 'doublingTimes')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.doublingTimes.DoublingTime_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Doubling Time (h)', 'FontWeight', 'bold');
                    ylabel('Doubling Time (h)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 6. Curve Fitting
                if isfield(analysisResults, 'curveFitting')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.curveFitting.R_squared);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Curve Fit Quality (R²)', 'FontWeight', 'bold');
                    ylabel('R²');
                    grid on;
                end
                
            else % Luciferase Analysis Bar Charts
                plotCount = 1;
                
                % 1. Peak Activity
                if isfield(analysisResults, 'peakActivity')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.peakActivity.PeakActivity);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Peak Activity', 'FontWeight', 'bold');
                    ylabel('Peak Value');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 2. Expression Rate
                if isfield(analysisResults, 'expressionRates')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.expressionRates.ExpressionRate_units_per_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Expression Rate', 'FontWeight', 'bold');
                    ylabel('Rate (units/h)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 3. Total Expression (AUC)
                if isfield(analysisResults, 'auc')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.auc.AUC_units_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Total Expression (AUC)', 'FontWeight', 'bold');
                    ylabel('AUC (units·h)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 4. Fold Change
                if isfield(analysisResults, 'foldChange')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.foldChange.FoldChange);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Fold Change', 'FontWeight', 'bold');
                    ylabel('Fold Change');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 5. Expression Duration
                if isfield(analysisResults, 'duration')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.duration.Duration_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Expression Duration (h)', 'FontWeight', 'bold');
                    ylabel('Duration (h)');
                    grid on;
                    plotCount = plotCount + 1;
                end
                
                % 6. Expression Efficiency
                if isfield(analysisResults, 'efficiency')
                    subplot(2, 3, plotCount);
                    b = bar(analysisResults.efficiency.Efficiency_units_per_h);
                    b.FaceColor = 'flat';
                    for i = 1:numGroups
                        b.CData(i,:) = groupColors(i,:);
                    end
                    set(gca, 'XTickLabel', groupNames);
                    title('Expression Efficiency', 'FontWeight', 'bold');
                    ylabel('Efficiency (units/h)');
                    grid on;
                end
            end
            
            sgtitle(sprintf('%s Analysis - Main Metrics Bar Charts', dataType), ...
                   'FontSize', 16, 'FontWeight', 'bold');
            
            % Save the main bar charts
            saveas(fig1, fullfile(outputPath, sprintf('%s_Bar_Charts_%s.png', dataType, timestamp)));
        end
        
        % PLOT 2: Time to Thresholds (7th analysis)
        if isfield(analysisResults, 'thresholds')
            fig2 = figure('Name', sprintf('%s Time to Thresholds - %s', dataType, timestamp), ...
                         'Position', [150, 150, 1000, 600]);
            
            thresholds = analysisResults.thresholds;
            numThresholds = length(thresholds);
            
            % Create subplot for each threshold
            for t = 1:numThresholds
                subplot(1, numThresholds, t);
                
                threshTable = thresholds{t};
                threshValues = threshTable{:, 2}; % Time values
                
                % Only plot groups that reached the threshold
                validGroups = ~isnan(threshValues) & threshValues > 0;
                
                if any(validGroups)
                    validTimes = threshValues(validGroups);
                    validNames = groupNames(validGroups);
                    validColors = groupColors(validGroups, :);
                    
                    b = bar(validTimes);
                    b.FaceColor = 'flat';
                    for i = 1:length(validTimes)
                        b.CData(i,:) = validColors(i,:);
                    end
                    
                    set(gca, 'XTickLabel', validNames);
                    
                    % Extract threshold value
                    varName = threshTable.Properties.VariableNames{2};
                    if strcmp(dataType, 'OD')
                        try
                            threshStr = extractBetween(varName, 'OD', '_h');
                            if ~isempty(threshStr)
                                title(sprintf('Time to OD %s', threshStr{1}), 'FontWeight', 'bold');
                            else
                                title(sprintf('Time to Threshold %d', t), 'FontWeight', 'bold');
                            end
                        catch
                            title(sprintf('Time to Threshold %d', t), 'FontWeight', 'bold');
                        end
                        ylabel('Time (hours)');
                    else
                        title(sprintf('Time to Threshold %d', t), 'FontWeight', 'bold');
                        ylabel('Time (hours)');
                    end
                    grid on;
                else
                    text(0.5, 0.5, 'No groups reached threshold', ...
                         'HorizontalAlignment', 'center', 'FontSize', 12);
                    title(sprintf('Threshold %d', t), 'FontWeight', 'bold');
                    set(gca, 'XLim', [0, 1], 'YLim', [0, 1]);
                end
            end
            
            sgtitle(sprintf('%s Time to Thresholds Analysis', dataType), ...
                   'FontSize', 16, 'FontWeight', 'bold');
            
            saveas(fig2, fullfile(outputPath, sprintf('%s_Thresholds_%s.png', dataType, timestamp)));
        end
        
        % PLOT 3: Statistical Comparisons (8th analysis)
        if numGroups >= 2
            fig3 = figure('Name', sprintf('%s Statistical Comparisons - %s', dataType, timestamp), ...
                         'Position', [200, 200, 800, 600]);
            
            % Create p-value matrix
            pMatrix = ones(numGroups, numGroups);
            groupData = analysisData.groupData;
            timeVectorInHours = analysisData.timeVectorInHours;
            
            % Calculate pairwise p-values
            for i = 1:numGroups
                for j = 1:numGroups
                    if i ~= j && ~isempty(groupData{i}) && ~isempty(groupData{j})
                        aucs_i = [];
                        aucs_j = [];
                        for k = 1:size(groupData{i}, 1)
                            aucs_i = [aucs_i, trapz(timeVectorInHours, groupData{i}(k, :))];
                        end
                        for k = 1:size(groupData{j}, 1)
                            aucs_j = [aucs_j, trapz(timeVectorInHours, groupData{j}(k, :))];
                        end
                        
                        try
                            [~, p_val] = ttest2(aucs_i, aucs_j);
                            pMatrix(i, j) = p_val;
                        catch
                            pMatrix(i, j) = 1;
                        end
                    end
                end
            end
            
            % Heatmap
            subplot(1, 2, 1);
            imagesc(pMatrix);
            colormap('hot');
            colorbar;
            caxis([0, 0.1]);
            
            % Add significance markers
            hold on;
            [rows, cols] = find(pMatrix < 0.05 & pMatrix > 0);
            for k = 1:length(rows)
                text(cols(k), rows(k), '*', 'HorizontalAlignment', 'center', ...
                     'FontSize', 20, 'Color', 'white', 'FontWeight', 'bold');
            end
            
            set(gca, 'XTick', 1:numGroups, 'XTickLabel', groupNames, ...
                     'YTick', 1:numGroups, 'YTickLabel', groupNames);
            title('P-values (AUC comparison)', 'FontWeight', 'bold');
            xlabel('Groups');
            ylabel('Groups');
            
            % Summary bar chart
            subplot(1, 2, 2);
            sigCounts = sum(pMatrix < 0.05, 2) - 1;
            
            b = bar(sigCounts);
            b.FaceColor = 'flat';
            for i = 1:numGroups
                b.CData(i,:) = groupColors(i,:);
            end
            
            set(gca, 'XTickLabel', groupNames);
            title('Significant Differences Count', 'FontWeight', 'bold');
            ylabel('Number of Significant Comparisons');
            grid on;
            
            sgtitle(sprintf('%s Statistical Comparisons (p < 0.05 marked with *)', dataType), ...
                   'FontSize', 16, 'FontWeight', 'bold');
            
            saveas(fig3, fullfile(outputPath, sprintf('%s_Statistics_%s.png', dataType, timestamp)));
        end
        
        msgbox(sprintf('All enhanced analysis plots generated!\n\nFiles created:\n• Main Metrics Bar Charts (6 analyses)\n• Time to Thresholds\n• Statistical Comparisons\n\nNote: Time-series plots created separately in basic analysis.'), 'All Plots Complete');
        
    catch ME
        msgbox(sprintf('Error generating plots:\n%s', ME.message), 'Plot Error');
        disp(getReport(ME, 'extended'));
    end
        end
        function debugFillData(groupAverage, groupStd, timeVectorInHours, groupIndex)
% Debug function to check data dimensions and content
    fprintf('\n=== DEBUG FILL DATA FOR GROUP %d ===\n', groupIndex);
    
    if ~isempty(groupAverage{groupIndex})
        avgData = groupAverage{groupIndex};
        stdData = groupStd{groupIndex};
        timeData = timeVectorInHours;
        
        fprintf('Original dimensions:\n');
        fprintf('  avgData: %dx%d\n', size(avgData, 1), size(avgData, 2));
        fprintf('  stdData: %dx%d\n', size(stdData, 1), size(stdData, 2));
        fprintf('  timeData: %dx%d\n', size(timeData, 1), size(timeData, 2));
        
        fprintf('Data ranges:\n');
        fprintf('  avgData: %.4f to %.4f\n', min(avgData), max(avgData));
        fprintf('  stdData: %.4f to %.4f\n', min(stdData), max(stdData));
        fprintf('  timeData: %.4f to %.4f\n', min(timeData), max(timeData));
        
        fprintf('Any NaN/Inf values:\n');
        fprintf('  avgData: %d NaN, %d Inf\n', sum(isnan(avgData)), sum(isinf(avgData)));
        fprintf('  stdData: %d NaN, %d Inf\n', sum(isnan(stdData)), sum(isinf(stdData)));
        fprintf('  timeData: %d NaN, %d Inf\n', sum(isnan(timeData)), sum(isinf(timeData)));
        
        % Test bounds calculation
        lowerBound = max(0, avgData - stdData);
        upperBound = avgData + stdData;
        
        fprintf('Bounds:\n');
        fprintf('  lowerBound: %.4f to %.4f\n', min(lowerBound), max(lowerBound));
        fprintf('  upperBound: %.4f to %.4f\n', min(upperBound), max(upperBound));
        fprintf('  any negative in lower: %d\n', sum(lowerBound < 0));
        
        fprintf('===============================\n\n');
    else
        fprintf('Group %d has empty data!\n', groupIndex);
    end
        end
        
        function success = simpleFillTest(ax, timeData, avgData, stdData, color)
% Simple fill test function - returns true if successful
    try
        % Force to row vectors
        timeData = timeData(:)';
        avgData = avgData(:)';
        stdData = stdData(:)';
        
        % Simple bounds
        lower = max(0, avgData - stdData);
        upper = avgData + stdData;
        
        % Create patch manually
        x_fill = [timeData, fliplr(timeData)];
        y_fill = [lower, fliplr(upper)];
        
        % Remove any NaN
        valid = isfinite(x_fill) & isfinite(y_fill);
        x_fill = x_fill(valid);
        y_fill = y_fill(valid);
        
        if length(x_fill) > 2
            if nargin < 1 || isempty(ax)
                patch(x_fill, y_fill, color, 'FaceAlpha', 0.3, 'EdgeColor', 'none');
            else
                patch(ax, x_fill, y_fill, color, 'FaceAlpha', 0.3, 'EdgeColor', 'none');
            end
            success = true;
        else
            success = false;
        end
    catch ME
        fprintf('Simple fill test failed: %s\n', ME.message);
        success = false;
    end
        end
        
    end
end
