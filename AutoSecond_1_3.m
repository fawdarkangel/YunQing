function varargout = AutoSecond_1_3(varargin)
% AUTOSECOND_1_3 MATLAB code for AutoSecond_1_3.fig
%      AUTOSECOND_1_3, by itself, creates a new AUTOSECOND_1_3 or raises the existing
%      singleton*.
%
%      H = AUTOSECOND_1_3 returns the handle to a new AUTOSECOND_1_3 or the handle to
%      the existing singleton*.
%
%      AUTOSECOND_1_3('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in AUTOSECOND_1_3.M with the given input arguments.
%
%      AUTOSECOND_1_3('Property','Value',...) creates a new AUTOSECOND_1_3 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before AutoSecond_1_3_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to AutoSecond_1_3_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help AutoSecond_1_3

% Last Modified by GUIDE v2.5 29-Mar-2018 07:54:36

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoSecond_1_3_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoSecond_1_3_OutputFcn, ...
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


% --- Executes just before AutoSecond_1_3 is made visible.
function AutoSecond_1_3_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to AutoSecond_1_3 (see VARARGIN)

% Choose default command line output for AutoSecond_1_3
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes AutoSecond_1_3 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = AutoSecond_1_3_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
winopen([cd,'\model\������.docx'])
