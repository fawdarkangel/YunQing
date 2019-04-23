function varargout = YunQing(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @YunQing_OpeningFcn, ...
                   'gui_OutputFcn',  @YunQing_OutputFcn, ...
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


% --- Executes just before YunQing is made visible.
function YunQing_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')
Cover = imread('Cover.png');
axes(handles.axes4);
imshow(Cover);
axis off

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);


% UIWAIT makes YunQing wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = YunQing_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;
% --- Executes during object creation, after setting all properties.
function axes4_CreateFcn(hObject, eventdata, handles)
%imshow(imread('Cover.png'));
%handles=guihandles;
%guidata(hObject,handles);
%Cover = imread('Cover.png');
%axes(handles.axes4);
%imshow(Cover);
%axis off


% --------------------------------------------------------------------
function Help_Callback(hObject, eventdata, handles)
% hObject    handle to Help (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function help1_Callback(hObject, eventdata, handles)
dos('help.txt');






% --------------------------------------------------------------------
function MenuFirst_Callback(hObject, eventdata, handles)
% hObject    handle to MenuFirst (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu1_Callback(hObject, eventdata, handles)
% hObject    handle to Menu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu1_1_Callback(hObject, eventdata, handles)
run DVDberichter;


% --------------------------------------------------------------------
function Menu2_Callback(hObject, eventdata, handles)
% hObject    handle to Menu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

function Menu2_1_Callback(hObject, eventdata, handles)
%open('Auto2_1.fig');
run Auto2_1;


% --------------------------------------------------------------------
function Menu3_Callback(hObject, eventdata, handles)
% hObject    handle to Menu3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu4_Callback(hObject, eventdata, handles)
% hObject    handle to Menu4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu5_Callback(hObject, eventdata, handles)
% hObject    handle to Menu5 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu6_Callback(hObject, eventdata, handles)
% hObject    handle to Menu6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu7_Callback(hObject, eventdata, handles)
% hObject    handle to Menu7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu6_1_Callback(hObject, eventdata, handles)
run Auto6_1;




% --------------------------------------------------------------------
function about_Callback(hObject, eventdata, handles)
dos('about.txt');


% --------------------------------------------------------------------


% --------------------------------------------------------------------
function MenuSecond_Callback(hObject, eventdata, handles)
% hObject    handle to MenuSecond (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function MenuSecond1_Callback(hObject, eventdata, handles)
% hObject    handle to MenuSecond1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menusecond1_1_Callback(hObject, eventdata, handles)
dos('D:\Autorepoter\DVDberichter.xls');


% --------------------------------------------------------------------
function Menu2_2_Callback(hObject, eventdata, handles)
run Auto2_2


% --------------------------------------------------------------------
function MenuSecond2_Callback(hObject, eventdata, handles)
% hObject    handle to MenuSecond2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function MenuSecond2_1_Callback(hObject, eventdata, handles)
dos('D:\Autorepoter\VWstossfaenger.xls');


% --------------------------------------------------------------------
function Menu7_1_Callback(hObject, eventdata, handles)
run Auto7_1;


% --------------------------------------------------------------------
function Menu2_3_Callback(hObject, eventdata, handles)
run Auto2_3;


% --------------------------------------------------------------------
function Menu_Six_Callback(hObject, eventdata, handles)
% hObject    handle to Menu_Six (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu_Six_1_Callback(hObject, eventdata, handles)
run AutoSix_1;


% --------------------------------------------------------------------
function Menu3_1_Callback(hObject, eventdata, handles)
% hObject    handle to Menu3_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu3_1_1_Callback(hObject, eventdata, handles)
run Auto3_1_1


% --------------------------------------------------------------------
function Menu3_1_2_Callback(hObject, eventdata, handles)
run Auto3_1_2


% --------------------------------------------------------------------
function Menu3_1_3_Callback(hObject, eventdata, handles)
run Auto3_1_3



% --------------------------------------------------------------------
function Untitled_3_Callback(hObject, eventdata, handles)
% hObject    handle to Untitled_3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu8_1_Callback(hObject, eventdata, handles)
run Auto8_1


% --------------------------------------------------------------------
function Menu_Second_Callback(hObject, eventdata, handles)
% hObject    handle to Menu_Second (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Meun_Second_1_Callback(hObject, eventdata, handles)
% hObject    handle to Meun_Second_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Meun_Second_2_Callback(hObject, eventdata, handles)
% hObject    handle to Meun_Second_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Meun_Second_1_1_Callback(hObject, eventdata, handles)
run AutoSecond_1_2;


% --------------------------------------------------------------------
function Meun_Second_1_2_Callback(hObject, eventdata, handles)
run AutoSecond_1_2;


% --------------------------------------------------------------------
function Meun_Second_1_3_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Meun_Second_1_4_Callback(hObject, eventdata, handles)
run Auto8_1;
% --------------------------------------------------------------------
function Menu4_1_Callback(hObject, eventdata, handles)
% --------------------------------------------------------------------
function Menu4_1_1_Callback(hObject, eventdata, handles)
run Auto4_1_1;
% --------------------------------------------------------------------
function Menu4_1_2_Callback(hObject, eventdata, handles)
run Auto4_1_2;
% --------------------------------------------------------------------
function Menu4_1_3_Callback(hObject, eventdata, handles)
run Auto4_1_3;


% --------------------------------------------------------------------
function Meun_Second_1_5_Callback(hObject, eventdata, handles)
run AutoSecond_1_5;


% --------------------------------------------------------------------
function Menu_Third_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu_Third_1_Callback(hObject, eventdata, handles)


% --------------------------------------------------------------------
function MenuThird_1_1_Callback(hObject, eventdata, handles)
run AutoThird_1_1;


% --------------------------------------------------------------------
function Menu1_2_Callback(hObject, eventdata, handles)
run Auto1_2;


% --------------------------------------------------------------------
function Menu3_2_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu3_2_1_Callback(hObject, eventdata, handles)
run Auto3_2_1


% --------------------------------------------------------------------
function Menu_Fourth_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu_Fourth_1_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu_Fourth_1_1_Callback(hObject, eventdata, handles)
run AutoFourth_1_1;


% --------------------------------------------------------------------
function Menu_Fifth_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu_Fifth_1_Callback(hObject, eventdata, handles)
run AutoFifth_1;


% --------------------------------------------------------------------
function Menu_Fourth_2_Callback(hObject, eventdata, handles)
% hObject    handle to Menu_Fourth_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function Menu_Fourth_2_1_Callback(hObject, eventdata, handles)
run AutoFourth_2_1;


% --------------------------------------------------------------------
function Menu3_3_Callback(hObject, eventdata, handles)


% --------------------------------------------------------------------
function Menu3_3_1_Callback(hObject, eventdata, handles)
run Auto3_3_1


% --------------------------------------------------------------------
function Menu7_2_Callback(hObject, eventdata, handles)
run Auto7_2


% --------------------------------------------------------------------
function Menu5_1_Callback(hObject, eventdata, handles)
run Auto5_1


% --------------------------------------------------------------------
function Menu6_2_Callback(hObject, eventdata, handles)
run Auto6_2


% --------------------------------------------------------------------
function Menu_Fourth_3_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu_Fourth_3_1_Callback(hObject, eventdata, handles)
run Auto6_1


% --------------------------------------------------------------------
function Menu6_3_Callback(hObject, eventdata, handles)
run Auto6_3


% --------------------------------------------------------------------
function Menu7_3_Callback(hObject, eventdata, handles)
run Auto7_3


% --------------------------------------------------------------------
function Menu2_4_Callback(hObject, eventdata, handles)
run Auto2_4

% --------------------------------------------------------------------
function Menu5_2_Callback(hObject, eventdata, handles)
run Auto5_2
