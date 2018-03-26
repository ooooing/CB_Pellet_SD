function varargout = PH_CB_PelletSD_20180323(varargin)
% PH_CB_PelletSD_20180323 MATLAB code for PH_CB_PelletSD_20180323.fig
%      PH_CB_PelletSD_20180323, by itself, creates a new PH_CB_PelletSD_20180323 or raises the existing
%      singleton*.
%
%      H = PH_CB_PelletSD_20180323 returns the handle to a new PH_CB_PelletSD_20180323 or the handle to
%      the existing singleton*.
%
%      PH_CB_PelletSD_20180323('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in PH_CB_PelletSD_20180323.M with the given input arguments.
%
%      PH_CB_PelletSD_20180323('Property','Value',...) creates a new PH_CB_PelletSD_20180323 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before PH_CB_PelletSD_20180323_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to PH_CB_PelletSD_20180323_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help PH_CB_PelletSD_20180323

% Last Modified by GUIDE v2.5 23-Mar-2018 17:55:27

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @PH_CB_PelletSD_20180323_OpeningFcn, ...
                   'gui_OutputFcn',  @PH_CB_PelletSD_20180323_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before PH_CB_PelletSD_20180323 is made visible.
function PH_CB_PelletSD_20180323_OpeningFcn(hObject, ~, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to PH_CB_PelletSD_20180323 (see VARARGIN)

% Choose default command line output for PH_CB_PelletSD_20180323
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);
global pathFMT

%pathFMT = 'C:\Program Files\MATLAB\MATLAB Runtime\';
pathFMT = 'E:\Image_Processing\Format\';
%pathFMT = 'X:\MATLAB\';
axes(handles.axes_Logo);imshow(sprintf('%s%s',pathFMT,'Logos.bmp'),'border','tight')


% UIWAIT makes PH_CB_PelletSD_20180323 wait for user response (see UIRESUME)
% uiwait(handles.figure1);

% --- Executes on button press in btn_Load.
function btn_Load_Callback(~, ~, handles) 
% hObject    handle to btn_Load (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% 1. Load

global file
global path
global Thrs
global ini
global cirThr
global ms
global rect

% Select images 
[file,path] = uigetfile('*.jpg','Select File(s)','multiselect','on');

if iscell(file) == 0
    file = {file};
end

% Pre-allocation
Thrs = zeros(1,size(file,2));

ini = 1;
cirThr = 0.4;
% ms = 8; % pixel
ms = 40; % um, 置煽 pelletmeter
% rect = [1348,910,1605,1449]; % 102_1226\
rect = [1516 874 1571 929]; % 20180320\101_0320\

% Set initial value of slider & its step
set(handles.slider_Thr,'Value',60);
set(handles.edit_Thr,'String',num2str(round(get(handles.slider_Thr,'Value'))));
set(handles.slider_Thr,'SliderStep',[1/255,5/255])

% Load first image
File = fullfile(path,file{1});
Original=imread(File);

axes(handles.axes1);
imshow(Original,'border','tight')
set(handles.edit_Status,'String','Loading Done');


% --- Outputs from this function are returned to the command line.
function varargout = PH_CB_PelletSD_20180323_OutputFcn(~, ~, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

% --- Executes on button press in btn_Load.
function btn_Analysis_Callback(~, ~, handles)
% hObject    handle to btn_Load (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% 5. Analysis

global Thrs
global path
global file
global pathFMT
global cirThr
global ms
global rect
global scaleFact

for x = 1:size(file,2)
    File = fullfile(path,file{x});
    Ori = imread(File);
    Ori = imcrop(Ori,rect);
    Ori = imsharpen(Ori,'Radius',2,'Amount',1); 
    hdint = vision.Deinterlacer;
    cOri = step(hdint,Ori);
    % Grayscale image
    Gr = rgb2gray(cOri);
    BW = im2bw(Gr,Thrs(x)/255);
    BW = imfill(BW,'holes');
    BW = imclearborder(BW,8);
    psb = str2double(get(handles.edit_Pixel,'String'));
    msb = str2double(get(handles.edit_Metric,'String'))*1000;
    mspix = ceil(((ms/2)^2*pi)*(psb/msb)^2);    
    BW = bwareaopen(BW,mspix);    
    [cc,N] = bwlabel(BW,8);
    stats = regionprops(cc,'Centroid','Area','PixelList','Perimeter');
    Center = [stats.Centroid];    
    Peri = [stats.Perimeter];
    Area = [stats.Area];
    Circularity = (4*pi*Area)./(Peri.^2);    
    for ii = 1:N
        if Circularity(ii) < 1-cirThr || Circularity(ii) > 1+cirThr
           for iii = 1:size(stats(ii).PixelList,1)
               BW(stats(ii).PixelList(iii,2),stats(ii).PixelList(iii,1)) = 0;
           end
        end        
    end    
    clear Area Centertmp Center stats cc N Circularity  
    [cc,N] = bwlabel(BW,8);
    stats = regionprops(cc,'EquivDiameter','Centroid');
    Dia = [stats.EquivDiameter];
    Center = [stats.Centroid];    
    for ii = 1:N
        Dia(ii) = Dia(ii)*(scaleFact(1)*(rect(2)+Center(2*ii))+scaleFact(2));
    end
    
    DiaEx = cell(size(Dia,2),2);
    
    for ii = 1:size(DiaEx,1)
        DiaEx{ii,1} = Dia(ii);
        DiaEx{ii,2} = sprintf('=4/3*pi()*((A%i/2)^3)/1000000000',ii+1);
    end           
   
    [~,name] = fileparts(file{x});
    imwrite(BW,sprintf('%s%s_BW.jpg',path,name))
    imwrite(Ori,sprintf('%s%s_Cropped.jpg',path,name))    
    xlswrite(sprintf('%s%s',pathFMT,'Format_CBPelletSD.xlsx'),DiaEx,1,'A2');
    xlswrite(sprintf('%s%s',pathFMT,'Format_CBPelletSD.xlsx'),Thrs(x),1,'E15');
    
       
    % Default varargin:
    sheet = 2;
    dimensions = [0,1,0,16.45,381.12/size(Ori,1)*size(Ori,2),381.12;0,1,381.12/size(Ori,1)*size(Ori,2)+22.69,16.45,381.12/size(Ori,1)*size(Ori,2),381.12];
    % [LinkToFile,SaveWithDocument,Left,Top,Width,Height]                                                         
    % Get handle to Excel COM Server
    Excel = actxserver('Excel.Application');
    % Add a Workbook to a new excel-file
    ResultFile = sprintf('%s%s',pathFMT,'Format_CBPelletSD.xlsx');
    Workbook = invoke(Excel.Workbooks,'Add', ResultFile);
    % Add a sheet if sheet = 'Add'
    if strcmp(sheet, 'Add')
       Workbook.Worksheets.Add([], Workbook.Worksheets.Item(Workbook.Worksheets.Count));
       sheet = Excel.Sheets.Count;
    end
    % Get a handle to Sheets and select Sheet No
    Sheets = Excel.ActiveWorkBook.Sheets;
    SheetNo = get(Sheets, 'Item', sheet);
    SheetNo.Activate;
    % Get a handle to Shapes for Sheet No n
    Shapes = SheetNo.Shapes;
    
    % Add image(s - adjacent to each other)
    Shapes.AddPicture(sprintf('%s%s%s',path,name,'_Cropped.jpg'),dimensions(1,1), dimensions(1,2),...
        dimensions(1,3), dimensions(1,4), dimensions(1,5), dimensions(1,6));    % Original image
    Shapes.AddPicture(sprintf('%s%s%s',path,name,'_BW.jpg') ,dimensions(2,1), dimensions(2,2),...
        dimensions(2,3), dimensions(2,4), dimensions(2,5), dimensions(2,6));    % Binary image

    %Save and Quite Excel file    
    invoke(Workbook, 'SaveAs', [path sprintf('%s%s',name,'.xlsx')]);
    invoke(Excel,'Quit');
    delete(Excel);
    if N ~= 0 
       blank = cell(N,1);
       xlswrite(sprintf('%s%s',pathFMT,'Format_CBPelletSD.xlsx'),blank,1,'A2');
    end
    clear blank Results  Size hist_dp
    delete(sprintf('%s%s_Cropped.jpg',path,name));
    delete(sprintf('%s%s_BW.jpg',path,name));
end
set(handles.edit_Status,'String','Export Done');

function edit_Thr_Callback(~, ~, ~)
% hObject    handle to edit_Thr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Thr as text
%        str2double(get(hObject,'String')) returns contents of edit_Thr as a double

% --- Executes during object creation, after setting all properties.
function edit_Thr_CreateFcn(hObject, ~, ~)
% hObject    handle to edit_Thr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


function edit_Pixel_Callback(~, ~, ~)
% hObject    handle to edit_Pixel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Pixel as text
%        str2double(get(hObject,'String')) returns contents of edit_Pixel as a double


% --- Executes during object creation, after setting all properties.
function edit_Pixel_CreateFcn(hObject, ~, ~)
% hObject    handle to edit_Pixel (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in btn_Summary.
function btn_Summary_Callback(~, ~, handles)
% hObject    handle to btn_Summary (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% 6. Summary

global pathFMT

[fileEx,pathEx] = uigetfile({'*.xlsx';'*.xls'},'Select File(s)','multiselect','on');

if iscell(fileEx) == 0
    fileEx = {fileEx};
end
raw = cell(1,size(fileEx,2));
count = zeros(1,size(fileEx,2));

for i = 1:size(fileEx,2)
    [~,~,raw{i}] = xlsread(sprintf('%s%s',pathEx,fileEx{i}),1);    
    count(i) = xlsread(sprintf('%s%s',pathEx,fileEx{i}),1,'D17');       
end
clear x

cumcount = [0,cumsum(count)];
diaSummary = zeros(cumcount(end),1);


for i = 1:size(cumcount,2)-1
    diaSummary(cumcount(i)+1:cumcount(i+1)) = [raw{i}{2:count(i)+1}]';
end

Results = {'Summary','';'Counts (ea)',size(diaSummary,1);'Mean (レm)',round(mean(diaSummary));'StDev (レm)',round(std(diaSummary));'Min (レm)',round(min(diaSummary));'Max (レm)',round(max(diaSummary));};
xlswrite(sprintf('%s%s',pathFMT,'Format_CBPelletSD_Integ.xlsx'),Results,1,'C16');
xlswrite(sprintf('%s%s',pathFMT,'Format_CBPelletSD_Integ.xlsx'),diaSummary,1,'A2');
summaryname = 'Summary';
% Default varargin:
sheet = 1;

% Get handle to Excel COM Server
Excel = actxserver('Excel.Application');
% Add a Workbook to a new excel-file
ResultFile = sprintf('%s%s',pathFMT,'Format_CBPelletSD_Integ.xlsx');
Workbook = invoke(Excel.Workbooks,'Add', ResultFile);
% Add a sheet if sheet = 'Add'
if strcmp(sheet, 'Add')
   Workbook.Worksheets.Add([], Workbook.Worksheets.Item(Workbook.Worksheets.Count));
   sheet = Excel.Sheets.Count;
end
% Get a handle to Sheets and select Sheet No
Sheets = Excel.ActiveWorkBook.Sheets;
SheetNo = get(Sheets, 'Item', sheet);
SheetNo.Activate;   

%Save and Quite Excel file    
invoke(Workbook, 'SaveAs', [pathEx sprintf('%s%s',summaryname,'.xlsx')]);
invoke(Excel,'Quit');
delete(Excel);
blank = cell(cumcount(end)+1,1);
xlswrite(sprintf('%s%s',pathFMT,'Format_CBPelletSD_Integ.xlsx'),blank,1,'A2');
clear blank 

set(handles.edit_Status,'String','Summary Done');

% --- Executes on slider movement.
function slider_Thr_Callback(~, ~, handles)
% hObject    handle to slider_Thr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'Value') returns position of slider
%        get(hObject,'Min') and get(hObject,'Max') to determine range of slider

global path
global file
global ini
global ms
global psb
global rect
global scaleFact
global cirThr

% 3. Threshold value

% Thresholding value load
Thr = round(get(handles.slider_Thr,'Value'));
set(handles.edit_Thr,'String',num2str(Thr));
% msb = str2double(get(handles.edit_Pixel,'String'));
% msb = 1;

% Load image
File = fullfile(path,file{ini});
Ori = imread(File);
Ori = imcrop(Ori,rect);
Ori = imsharpen(Ori,'Radius',2,'Amount',1); 
hdint = vision.Deinterlacer;
cOri = step(hdint,Ori);
% Grayscale image
Gr = rgb2gray(cOri);
% Binary image
BW = im2bw(Gr,Thr/255);
psb = 1;
BW = imfill(BW,'holes');
BW = imclearborder(BW,8);
[cc,N] = bwlabel(BW,8);
stats = regionprops(cc,'Centroid','Area','PixelList','Perimeter');
Center = [stats.Centroid];    
Peri = [stats.Perimeter];
Area = [stats.Area];
Circularity = (4*pi*Area)./(Peri.^2);    
for ii = 1:N
    mstmp = ((ms/2)^2*pi)/(scaleFact(1)*(rect(2)+Center(2*ii))+scaleFact(2))^2;
    if Area(ii) < mstmp || Circularity(ii) < 1-cirThr || Circularity(ii) > 1+cirThr
       for iii = 1:size(stats(ii).PixelList,1)
           BW(stats(ii).PixelList(iii,2),stats(ii).PixelList(iii,1)) = 0;
       end
    end        
end    
clear Area Centertmp Center stats cc N Circularity  

axes(handles.axes1)
imshow(Gr,'border','tight')
axes(handles.axes2)
imshow(BW,'border','tight')



% --- Executes during object creation, after setting all properties.
function slider_Thr_CreateFcn(hObject, ~, ~)
% hObject    handle to slider_Thr (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end


% --- Executes on button press in btn_ThrSet.
function btn_ThrSet_Callback(~, ~, handles)
% hObject    handle to btn_ThrSet (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

global Thrs
global ini
global file
global path
global msb
global psb
global ms
global psbs
global mss

if ini ~= size(file,2)
    msb(ini) = str2double(get(handles.edit_Pixel,'String'));
    Thrs(ini) = round(get(handles.slider_Thr,'Value'));
    mss(ini) = ms;
    psbs(ini) = psb;
    ini = ini+1;
    File = fullfile(path,file{ini});
    Original = imread(File);
    axes(handles.axes1)
    imshow(Original,'border','tight')
else
    Thrs(ini) = get(handles.slider_Thr,'Value');
    msb(ini) = str2double(get(handles.edit_Pixel,'String'));
    mss(ini) = ms;
    psbs(ini) = psb;
    set(handles.edit_Status,'String','Thresholding Done');
end    
    

% --- Executes during object creation, after setting all properties.
function edit_Status_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_Status (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% handles    empty - handles not created until after all CreateFcns called

% --- Executes during object creation, after setting all properties.
function edit_Metric_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit_Metric (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit_Metric_Callback(hObject, eventdata, handles)
% hObject    handle to edit_Metric (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Metric as text
%        str2double(get(hObject,'String')) returns contents of edit_Metric as a double



function edit_Status_Callback(hObject, eventdata, handles)
% hObject    handle to edit_Status (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit_Status as text
%        str2double(get(hObject,'String')) returns contents of edit_Status as a double



function edit10_Callback(hObject, eventdata, handles)
% hObject    handle to edit10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit10 as text
%        str2double(get(hObject,'String')) returns contents of edit10 as a double


% --- Executes during object creation, after setting all properties.
function edit10_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit10 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on slider movement.
function slider_Cir_Callback(hObject, eventdata, handles)
% hObject    handle to slider_Cir (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'Value') returns position of slider
%        get(hObject,'Min') and get(hObject,'Max') to determine range of slider


% --- Executes during object creation, after setting all properties.
function slider_Cir_CreateFcn(hObject, eventdata, handles)
% hObject    handle to slider_Cir (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end


% --- Executes on button press in btn_SetCir.
function btn_SetCir_Callback(hObject, eventdata, handles)
% hObject    handle to btn_SetCir (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
