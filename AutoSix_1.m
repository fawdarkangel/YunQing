function varargout = AutoSix_1(varargin)
% AUTOSIX_1 MATLAB code for AutoSix_1.fig
%      AUTOSIX_1, by itself, creates a new AUTOSIX_1 or raises the existing
%      singleton*.
%
%      H = AUTOSIX_1 returns the handle to a new AUTOSIX_1 or the handle to
%      the existing singleton*.
%
%      AUTOSIX_1('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in AUTOSIX_1.M with the given input arguments.
%
%      AUTOSIX_1('Property','Value',...) creates a new AUTOSIX_1 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before AutoSix_1_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to AutoSix_1_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help AutoSix_1

% Last Modified by GUIDE v2.5 09-Apr-2018 12:20:56

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoSix_1_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoSix_1_OutputFcn, ...
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


% --- Executes just before AutoSix_1 is made visible.
function AutoSix_1_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to AutoSix_1 (see VARARGIN)

% Choose default command line output for AutoSix_1
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes AutoSix_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = AutoSix_1_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
b1=47.244/1.25238*str2num(get(handles.edit1,'String'));
b2=47.244/1.25238*str2num(get(handles.edit2,'String'));
[filename,pathname,fileindex]=uigetfile('*.bmp;*.png;*.jpg','请选择图片','MultiSelect','on');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    
elseif ~iscell(filename);
    if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
              
    end
    Fileadress=strcat(pathname,'result\');
Filename{1}=strcat(pathname,filename);
t1=waitbar(0,'正在处理');
   fig{1}=imresize(imread(Filename{1}),[b2 b1]);
         sfilename=[Fileadress,filename];
    waitbar(100);
         imwrite(fig{1},sfilename);
close(t1);
winopen(Fileadress);
elseif length(filename)>1
    if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
end
   Fileadress=strcat(pathname,'result\');
   t1=waitbar(0,'正在处理');
     for i=1:length(filename)
         Filename{i}=strcat(pathname,filename{i});
     end
      for i=1:length(filename)
         fig{1,i}=imresize(imread(Filename{i}),[b2 b1]);
                 sfilename=[Fileadress,filename{i}];
    waitbar(i/length(filename));
         imwrite(fig{1,i},sfilename);
      end
      close(t1);
      winopen(Fileadress)
end


   
   
   
   
    
    
