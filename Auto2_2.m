function varargout = Auto2_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_2_OutputFcn, ...
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


% --- Executes just before Auto2_2 is made visible.
function Auto2_2_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);

axes(handles.axes1);

axis off
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto2_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto2_2_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



function edit1_Callback(hObject, eventdata, handles)

function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
cla(handles.axes1);
val1=get(handles.popupmenu1,'Value');
axes(handles.axes1);
if val1==2
Cover = imread('Auto2_2STF vorne.png');
imshow(Cover);
elseif val1==3||val1==4
  Cover = imread('Auto2_2KSG.png');
imshow(Cover);
elseif val1==5||val1==6
  Cover = imread('Auto2_2LueGi.png');
imshow(Cover);
elseif val1==7
     Cover = imread('Auto2_2STF hinten.png');
imshow(Cover);
end


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

handles=guihandles;
guidata(hObject,handles);
if isempty(get(handles.edit1,'String'))
    msgbox('请输入车型名称');
    return;
end 
list=get(handles.popupmenu1,'String');
val1=get(handles.popupmenu1,'Value');
%%创建统计文件输出文件夹
Fileaddress1=char('D:\Autorepoter\VWstossfaenger.xls');
  if ~exist('D:\Autorepoter','dir')
      mkdir('D:\Autorepoter');
  end
  if ~exist([Fileaddress1]) %创建统计文件
      xlswrite([Fileaddress1],{'车型'},'Sheet1','A1');
       xlswrite([Fileaddress1],{'Punkt'},'Sheet1','B1');
       xlswrite([Fileaddress1],{'Kraft（N）'},'Sheet1','C1');
       xlswrite([Fileaddress1],{'Weg（mm）'},'Sheet1','D1');
       xlswrite([Fileaddress1],{'刚度'},'Sheet1','E1');
       xlswrite([Fileaddress1],{'部位'},'Sheet1','F1');
       xlswrite([Fileaddress1],{'日期'},'Sheet1','G1');
  end
    [num text alldata]=xlsread('D:\Autorepoter\VWstossfaenger.xls');
            SZ=size(alldata,1);%SZ为当前工作表行数

   if get(handles.checkbox2,'Value')==1
                [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据');
                if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
                    msgbox('导入文件失败');
                else
                    t1=waitbar(0,'正在读入数据');
                    Filename=strcat(pathname,filename);
                    [Type Sheet Format]=xlsfinfo(Filename) ;
                    for i=1:length(Sheet)
                        MP{i}=xlsread(Filename,char(Sheet{1,i}));
                        waitbar(i/(length(Sheet)));
                    end
                end
   else
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
else  
    t1=waitbar(0,'正在读入数据');        

     for i=1:length(filename)
         Filename{i}=strcat(pathname,filename{i});
         [Type Sheet Format]=xlsfinfo(Filename{i}) ;
         sheet{i}=Sheet;
     end
      for i=1:length(filename)
      MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
       try
       system('taskkill/IM excel.exe');
       end
   waitbar(i/length(filename));
      end
   
end
   end

if get(handles.checkbox1,'Value')==1
    for i=1:length(MP)
    MP_MIDDLE{1,i}(:,1)=MP{1,i}(:,2);
    MP_MIDDLE{1,i}(:,2)=MP{1,i}(:,1);
    end
    MP=MP_MIDDLE;
end


    


close(t1);
zihao=9;
t2=waitbar(0,'正在生成图片');
if val1==2
  if ~exist('pathname\STFvorne','dir')
      mkdir(pathname,'STFvorne');
  end
  
   file_usr=strcat(cd,'\model\Auto2_2Stossfaenger\Auto2_2_300.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'STFvorne\STFvorne.pptx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
 
 try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 

 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end
  
  
  
   Lstfvorne=strcat(pathname,'STFvorne\');%合成保存图片路径Ｌ
   Fileaddress=char(strcat(Lstfvorne,'STFvorne.xls'));
for i=1:length(MP)
     h=figure(i);
     plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
set(h,'visible','off');
 set(h,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
        set(h,'color','w')
        set(gca,'FontSize',zihao);
            xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);        
            title(['MP',num2str(i)],'FontSize',zihao);
            axis([0 inf 0 inf]);
    sfilename=[Lstfvorne,'MP' num2str(i) '.jpg'];
    if max(MP{1,i}(:,2))<400
        F50=find(MP{1,i}(:,2)>=50,1,'First');
   F100=find(MP{1,i}(:,2)>=100,1,'First');
  F150=find(MP{1,i}(:,2)>=150,1,'First');
  F200=find(MP{1,i}(:,2)>=200,1,'First');
  F250=find(MP{1,i}(:,2)>=250,1,'First');
  F300=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F50,2);F(1,2)=MP{1,i}(F50,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F100,2);F(2,2)=MP{1,i}(F100,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F150,2);F(3,2)=MP{1,i}(F150,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F200,2);F(4,2)=MP{1,i}(F200,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F250,2);F(5,2)=MP{1,i}(F250,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F300,2);F(6,2)=MP{1,i}(F300,1);F(6,3)=F(6,1)/F(6,2);
kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3); %记录最大变形及刚度
    elseif  max(MP{1,i}(:,2))>400
     F83=find(MP{1,i}(:,2)>=83,1,'First');
   F166=find(MP{1,i}(:,2)>=166,1,'First');
  F249=find(MP{1,i}(:,2)>=249,1,'First');
  F332=find(MP{1,i}(:,2)>=332,1,'First');
  F415=find(MP{1,i}(:,2)>=415,1,'First');
  F500=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F83,2);F(1,2)=MP{1,i}(F83,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F166,2);F(2,2)=MP{1,i}(F166,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F249,2);F(3,2)=MP{1,i}(F249,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F332,2);F(4,2)=MP{1,i}(F332,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F415,2);F(5,2)=MP{1,i}(F415,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F500,2);F(6,2)=MP{1,i}(F500,1);F(6,3)=F(6,1)/F(6,2);
   kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3);%记录最大变形及刚度
    end
    f=getframe(h);
           imwrite(f.cdata,sfilename);
          %saveas(h,sfilename);

                xlswrite(Fileaddress,{'Kraft(N)'},[strcat('Sheet',num2str(i))],'A1');
      xlswrite(Fileaddress,{'Weg(mm)'},[strcat('Sheet',num2str(i))],'B1');
      xlswrite(Fileaddress,{'Spez.Wert(N/mm)'},[strcat('Sheet',num2str(i))],'C1');
  xlswrite(Fileaddress,F,[strcat('Sheet',num2str(i))],'A2');

close(h);
         Slidesn=Slides.Item(i);
 for m=1:6
     for n=1:3
         Slidesn.Shapes.Range.Item(3).Table.Cell(m+1,n).Shape.TextFrame.TextRange.Text=num2str(F(m,n),'%.0f');
      end
 end
 Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800)
        Slidesn.Shapes.Range.Item(2).Table.Cell(2,2).Shape.TextFrame.TextRange.Text=['MP',num2str(i),' bei RT'];
       Slidesn.Shapes.Range.Item(2).Table.Cell(1,2).Shape.TextFrame.TextRange.Text= get(handles.edit1,'String');
waitbar(i/(length(MP)+1));
end

Presentation.Save
%system('taskkill/IM POWERPNT.exe');


  p=find(kraft_weg(:,3)==min(kraft_weg(:,3)));
    MPwegmin=strcat('MP',num2str(p));
       Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));
     Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));
     Ezuobiao=strcat('E',num2str(SZ+1));Fzuobiao=strcat('F',num2str(SZ+1)); Gzuobiao=strcat('G',num2str(SZ+1));                               %生成写入EXCEL单元坐标
   xlswrite([Fileaddress1],{get(handles.edit1,'String')},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress1],{MPwegmin},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress1],kraft_weg(p,1),'Sheet1',[Czuobiao]);%写入C列对应变形
      xlswrite([Fileaddress1],kraft_weg(p,2),'Sheet1',[Dzuobiao]);%写入D列对应力
      xlswrite([Fileaddress1],kraft_weg(p,3),'Sheet1',[Ezuobiao]);%写入E列最小刚度
   xlswrite([Fileaddress1],{list{val1}},'Sheet1',[Fzuobiao]);%写入D列温度
   xlswrite([Fileaddress1],{date},'Sheet1',[Gzuobiao]);%写入E列时间
    winopen(Lstfvorne);

    
 %%计算格栅X方向   
elseif val1==3
     if ~exist('pathname\KSGx','dir')
      mkdir(pathname,'KSGx');
     end
  
   Lksgx=strcat(pathname,'KSGx\');%合成保存图片路径Ｌ
          Fileaddress=char(strcat(Lksgx,'KSGx.xls'));
          
           file_usr=strcat(cd,'\model\Auto2_2Stossfaenger\Auto2_2_100_KSGX.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'KSGx\KSGx.pptx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
 
 try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 

 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end
          
          
          
          
for i=1:length(MP)
     h=figure(i);
     plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
     set(h,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
        set(h,'color','w')
        set(gca,'FontSize',zihao);
set(h,'visible','off');
            xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);title(['MP',num2str(i)],'FontSize',zihao);
            axis([0 inf 0 110]);
  %%获取各个力位置坐标
  F15=find(MP{1,i}(:,2)>=15,1,'First');
   F30=find(MP{1,i}(:,2)>=30,1,'First');
  F45=find(MP{1,i}(:,2)>=45,1,'First');
  F60=find(MP{1,i}(:,2)>=60,1,'First');
  F80=find(MP{1,i}(:,2)>=80,1,'First');
  F100=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F15,2);F(1,2)=MP{1,i}(F15,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F30,2);F(2,2)=MP{1,i}(F30,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F45,2);F(3,2)=MP{1,i}(F45,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F60,2);F(4,2)=MP{1,i}(F60,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F80,2);F(5,2)=MP{1,i}(F80,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F100,2);F(6,2)=MP{1,i}(F100,1);F(6,3)=F(6,1)/F(6,2);
 kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3);%记录最大变形及刚度
    xlswrite(Fileaddress,{'Kraft(N)'},[strcat('Sheet',num2str(i))],'A1');
      xlswrite(Fileaddress,{'Weg(mm)'},[strcat('Sheet',num2str(i))],'B1');
      xlswrite(Fileaddress,{'Spez.Wert(N/mm)'},[strcat('Sheet',num2str(i))],'C1');
  xlswrite(Fileaddress,F,[strcat('Sheet',num2str(i))],'A2');
      sfilename=[Lksgx,'MP' num2str(i) '.jpg'];
           f=getframe(h);
           imwrite(f.cdata,sfilename);
           
           
close(h);
         Slidesn=Slides.Item(i);
 for m=1:6
     for n=1:3
         Slidesn.Shapes.Range.Item(3).Table.Cell(m+1,n).Shape.TextFrame.TextRange.Text=num2str(F(m,n),'%.0f');
                  end
 end
 Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800);
        Slidesn.Shapes.Range.Item(2).Table.Cell(2,2).Shape.TextFrame.TextRange.Text=['MP',num2str(i),' bei RT'];
       Slidesn.Shapes.Range.Item(2).Table.Cell(1,2).Shape.TextFrame.TextRange.Text= get(handles.edit1,'String');

   waitbar(i/(length(MP)+1));
              
end
Presentation.Save
%system('taskkill/IM POWERPNT.exe');
           
           
       

    p=find(kraft_weg(:,3)==min(kraft_weg(:,3)));
    MPwegmin=strcat('MP',num2str(p));
       Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));
     Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));
     Ezuobiao=strcat('E',num2str(SZ+1));Fzuobiao=strcat('F',num2str(SZ+1)); Gzuobiao=strcat('G',num2str(SZ+1));                               %生成写入EXCEL单元坐标
   xlswrite([Fileaddress1],{get(handles.edit1,'String')},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress1],{MPwegmin},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress1],kraft_weg(p,1),'Sheet1',[Czuobiao]);%写入C列对应变形
      xlswrite([Fileaddress1],kraft_weg(p,2),'Sheet1',[Dzuobiao]);%写入D列对应力
      xlswrite([Fileaddress1],kraft_weg(p,3),'Sheet1',[Ezuobiao]);%写入E列最小刚度
   xlswrite([Fileaddress1],{list{val1}},'Sheet1',[Fzuobiao]);%写入D列温度
   xlswrite([Fileaddress1],{date},'Sheet1',[Gzuobiao]);%写入E列时间

     winopen(Lksgx);
   
   
  %%计算格栅Z方向    
  elseif val1==4
     if ~exist('pathname\KSGz','dir')
      mkdir(pathname,'KSGz');
     end
  
     
   Lksgz=strcat(pathname,'KSGz\');%合成保存图片路径Ｌ 
   Fileaddress=char(strcat(Lksgz,'KSGz.xls'));
   
    file_usr=strcat(cd,'\model\Auto2_2Stossfaenger\Auto2_2_100_KSGZ.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'KSGz\KSGz.pptx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
 
 try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 

 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end
   
   
   
   
for i=1:length(MP)
     h=figure(i);
     plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
set(h,'visible','off');
set(h,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
        set(h,'color','w')
        set(gca,'FontSize',zihao);
            xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);title(['MP',num2str(i)],'FontSize',zihao);
            axis([0 inf 0 110]);
  %%获取各个力位置坐标
  F15=find(MP{1,i}(:,2)>=15,1,'First');
   F30=find(MP{1,i}(:,2)>=30,1,'First');
  F45=find(MP{1,i}(:,2)>=45,1,'First');
  F60=find(MP{1,i}(:,2)>=60,1,'First');
  F80=find(MP{1,i}(:,2)>=80,1,'First');
  F100=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F15,2);F(1,2)=MP{1,i}(F15,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F30,2);F(2,2)=MP{1,i}(F30,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F45,2);F(3,2)=MP{1,i}(F45,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F60,2);F(4,2)=MP{1,i}(F60,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F80,2);F(5,2)=MP{1,i}(F80,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F100,2);F(6,2)=MP{1,i}(F100,1);F(6,3)=F(6,1)/F(6,2);
 kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3);%记录最大变形及刚度
    xlswrite(Fileaddress,{'Kraft(N)'},[strcat('Sheet',num2str(i))],'A1');
      xlswrite(Fileaddress,{'Weg(mm)'},[strcat('Sheet',num2str(i))],'B1');
      xlswrite(Fileaddress,{'Spez.Wert(N/mm)'},[strcat('Sheet',num2str(i))],'C1');
  xlswrite(Fileaddress,F,[strcat('Sheet',num2str(i))],'A2');
      sfilename=[Lksgz,'MP' num2str(i) '.jpg'];
           f=getframe(h);
           imwrite(f.cdata,sfilename); 
           close(h);
         Slidesn=Slides.Item(i);
 for m=1:6
     for n=1:3
         Slidesn.Shapes.Range.Item(3).Table.Cell(m+1,n).Shape.TextFrame.TextRange.Text=num2str(F(m,n),'%.0f');
        
         end
 end
  Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800);
        Slidesn.Shapes.Range.Item(2).Table.Cell(2,2).Shape.TextFrame.TextRange.Text=['MP',num2str(i),' bei RT'];
       Slidesn.Shapes.Range.Item(2).Table.Cell(1,2).Shape.TextFrame.TextRange.Text= get(handles.edit1,'String');
   waitbar(i/(length(MP)+1));
end

Presentation.Save
%system('taskkill/IM POWERPNT.exe');

           
           
       
p=find(kraft_weg(:,3)==min(kraft_weg(:,3)));
    MPwegmin=strcat('MP',num2str(p));
       Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));
     Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));
     Ezuobiao=strcat('E',num2str(SZ+1));Fzuobiao=strcat('F',num2str(SZ+1)); Gzuobiao=strcat('G',num2str(SZ+1));                               %生成写入EXCEL单元坐标
   xlswrite([Fileaddress1],{get(handles.edit1,'String')},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress1],{MPwegmin},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress1],kraft_weg(p,1),'Sheet1',[Czuobiao]);%写入C列对应变形
      xlswrite([Fileaddress1],kraft_weg(p,2),'Sheet1',[Dzuobiao]);%写入D列对应力
      xlswrite([Fileaddress1],kraft_weg(p,3),'Sheet1',[Ezuobiao]);%写入E列最小刚度
   xlswrite([Fileaddress1],{list{val1}},'Sheet1',[Fzuobiao]);%写入D列温度
   xlswrite([Fileaddress1],{date},'Sheet1',[Gzuobiao]);%写入E列时间
  
   winopen(Lksgz);
   
   
  %%计算LueGix方向     
   elseif val1==5
     if ~exist('pathname\LueGix','dir')
      mkdir(pathname,'LueGix');
     end
   Lluegix=strcat(pathname,'LueGix\');%合成保存图片路径Ｌ 
    Fileaddress=char(strcat(Lluegix,'LueGix.xls'));
    
     file_usr=strcat(cd,'\model\Auto2_2Stossfaenger\Auto2_2_100_LuegiX.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'LueGix\LueGix.pptx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
 
 try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 

 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end

    
    
for i=1:length(MP)
     h=figure(i);
     plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
set(h,'visible','off');
set(h,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
        set(h,'color','w')
        set(gca,'FontSize',zihao);
            xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);title(['MP',num2str(i)],'FontSize',zihao);
            axis([0 inf 0 110]);
  %%获取各个力位置坐标
  F15=find(MP{1,i}(:,2)>=15,1,'First');
   F30=find(MP{1,i}(:,2)>=30,1,'First');
  F45=find(MP{1,i}(:,2)>=45,1,'First');
  F60=find(MP{1,i}(:,2)>=60,1,'First');
  F80=find(MP{1,i}(:,2)>=80,1,'First');
  F100=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F15,2);F(1,2)=MP{1,i}(F15,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F30,2);F(2,2)=MP{1,i}(F30,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F45,2);F(3,2)=MP{1,i}(F45,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F60,2);F(4,2)=MP{1,i}(F60,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F80,2);F(5,2)=MP{1,i}(F80,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F100,2);F(6,2)=MP{1,i}(F100,1);F(6,3)=F(6,1)/F(6,2);
 kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3);%记录最大变形及刚度
    xlswrite(Fileaddress,{'Kraft(N)'},[strcat('Sheet',num2str(i))],'A1');
      xlswrite(Fileaddress,{'Weg(mm)'},[strcat('Sheet',num2str(i))],'B1');
      xlswrite(Fileaddress,{'Spez.Wert(N/mm)'},[strcat('Sheet',num2str(i))],'C1');
  xlswrite(Fileaddress,F,[strcat('Sheet',num2str(i))],'A2');
      sfilename=[Lluegix,'MP' num2str(i) '.jpg'];
           f=getframe(h);
           imwrite(f.cdata,sfilename);
           
close(h);
         Slidesn=Slides.Item(i);
 for m=1:6
     for n=1:3
         Slidesn.Shapes.Range.Item(3).Table.Cell(m+1,n).Shape.TextFrame.TextRange.Text=num2str(F(m,n),'%.0f');
         
         end
 end
 Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800);
        Slidesn.Shapes.Range.Item(2).Table.Cell(2,2).Shape.TextFrame.TextRange.Text=['MP',num2str(i),' bei RT'];
       Slidesn.Shapes.Range.Item(2).Table.Cell(1,2).Shape.TextFrame.TextRange.Text= get(handles.edit1,'String');
 waitbar(i/(length(MP)+1));
end

Presentation.Save
%system('taskkill/IM POWERPNT.exe');
           
           

p=find(kraft_weg(:,3)==min(kraft_weg(:,3)));
    MPwegmin=strcat('MP',num2str(p));
       Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));
     Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));
     Ezuobiao=strcat('E',num2str(SZ+1));Fzuobiao=strcat('F',num2str(SZ+1)); Gzuobiao=strcat('G',num2str(SZ+1));                               %生成写入EXCEL单元坐标
   xlswrite([Fileaddress1],{get(handles.edit1,'String')},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress1],{MPwegmin},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress1],kraft_weg(p,1),'Sheet1',[Czuobiao]);%写入C列对应变形
      xlswrite([Fileaddress1],kraft_weg(p,2),'Sheet1',[Dzuobiao]);%写入D列对应力
      xlswrite([Fileaddress1],kraft_weg(p,3),'Sheet1',[Ezuobiao]);%写入E列最小刚度
   xlswrite([Fileaddress1],{list{val1}},'Sheet1',[Fzuobiao]);%写入D列温度
   xlswrite([Fileaddress1],{date},'Sheet1',[Gzuobiao]);%写入E列时间
  
     winopen(Lluegix);
   
  
   
   %%计算进气格栅y
      elseif val1==6
     if ~exist('pathname\LueGiz','dir')
      mkdir(pathname,'LueGiz');
  end
   Lluegiy=strcat(pathname,'LueGiz\');%合成保存图片路径Ｌ 
    Fileaddress=char(strcat(Lluegiy,'LueGiy.xls'));
     file_usr=strcat(cd,'\model\Auto2_2Stossfaenger\Auto2_2_100_LuegiZ.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'LueGiz\LueGiz.pptx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
 
 try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 

 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end

    
    
for i=1:length(MP)
     h=figure(i);
     plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
set(h,'visible','off');
set(h,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
        set(h,'color','w')
        set(gca,'FontSize',zihao);
            xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);title(['MP',num2str(i)],'FontSize',zihao);
            axis([0 inf 0 110]);
  %%获取各个力位置坐标
  F15=find(MP{1,i}(:,2)>=15,1,'First');
   F30=find(MP{1,i}(:,2)>=30,1,'First');
  F45=find(MP{1,i}(:,2)>=45,1,'First');
  F60=find(MP{1,i}(:,2)>=60,1,'First');
  F80=find(MP{1,i}(:,2)>=80,1,'First');
  F100=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F15,2);F(1,2)=MP{1,i}(F15,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F30,2);F(2,2)=MP{1,i}(F30,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F45,2);F(3,2)=MP{1,i}(F45,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F60,2);F(4,2)=MP{1,i}(F60,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F80,2);F(5,2)=MP{1,i}(F80,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F100,2);F(6,2)=MP{1,i}(F100,1);F(6,3)=F(6,1)/F(6,2);
 kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3);%记录最大变形及刚度
    xlswrite(Fileaddress,{'Kraft(N)'},[strcat('Sheet',num2str(i))],'A1');
      xlswrite(Fileaddress,{'Weg(mm)'},[strcat('Sheet',num2str(i))],'B1');
      xlswrite(Fileaddress,{'Spez.Wert(N/mm)'},[strcat('Sheet',num2str(i))],'C1');
  xlswrite(Fileaddress,F,[strcat('Sheet',num2str(i))],'A2');
      sfilename=[Lluegiy,'MP' num2str(i) '.jpg'];
           f=getframe(h);
           imwrite(f.cdata,sfilename); 
           
close(h);
         Slidesn=Slides.Item(i);
 for m=1:6
     for n=1:3
         Slidesn.Shapes.Range.Item(3).Table.Cell(m+1,n).Shape.TextFrame.TextRange.Text=num2str(F(m,n),'%.0f');
         
         end
 end
 Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800);
        Slidesn.Shapes.Range.Item(2).Table.Cell(2,2).Shape.TextFrame.TextRange.Text=['MP',num2str(i),' bei RT'];
       Slidesn.Shapes.Range.Item(2).Table.Cell(1,2).Shape.TextFrame.TextRange.Text= get(handles.edit1,'String');
  waitbar(i/(length(MP)+1));
end   

Presentation.Save
%system('taskkill/IM POWERPNT.exe');

           
           
           
       
p=find(kraft_weg(:,3)==min(kraft_weg(:,3)));
    MPwegmin=strcat('MP',num2str(p));
       Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));
     Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));
     Ezuobiao=strcat('E',num2str(SZ+1));Fzuobiao=strcat('F',num2str(SZ+1)); Gzuobiao=strcat('G',num2str(SZ+1));                               %生成写入EXCEL单元坐标
   xlswrite([Fileaddress1],{get(handles.edit1,'String')},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress1],{MPwegmin},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress1],kraft_weg(p,1),'Sheet1',[Czuobiao]);%写入C列对应变形
      xlswrite([Fileaddress1],kraft_weg(p,2),'Sheet1',[Dzuobiao]);%写入D列对应力
      xlswrite([Fileaddress1],kraft_weg(p,3),'Sheet1',[Ezuobiao]);%写入E列最小刚度
   xlswrite([Fileaddress1],{list{val1}},'Sheet1',[Fzuobiao]);%写入D列温度
   xlswrite([Fileaddress1],{date},'Sheet1',[Gzuobiao]);%写入E列时间
 
    winopen(Lluegiy);
   
   
   
   
   %%计算后杠
   elseif val1==7
     if ~exist('pathname\STFhinten','dir')
      mkdir(pathname,'STFhinten');
     end
       file_usr=strcat(cd,'\model\Auto2_2Stossfaenger\Auto2_2_300.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'STFhinten\STFhinten.pptx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
 
 try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 

 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end
 
  
   Lstfhinten=strcat(pathname,'STFhinten\');%合成保存图片路径Ｌ 
   Fileaddress=char(strcat(Lstfhinten,'STFhinten.xls'));
for i=1:length(MP)
     h=figure(i);
     plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
set(h,'visible','off');
set(h,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
        set(h,'color','w')
        set(gca,'FontSize',zihao);
            xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);title(['MP',num2str(i)],'FontSize',zihao);
            axis([0 inf 0 310]);
  %%获取各个力位置坐标
  F50=find(MP{1,i}(:,2)>=50,1,'First');
   F100=find(MP{1,i}(:,2)>=100,1,'First');
  F150=find(MP{1,i}(:,2)>=150,1,'First');
  F200=find(MP{1,i}(:,2)>=200,1,'First');
  F250=find(MP{1,i}(:,2)>=250,1,'First');
  F300=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)),1,'First');
  F(1,1)=MP{1,i}(F50,2);F(1,2)=MP{1,i}(F50,1);F(1,3)=F(1,1)/F(1,2);
  F(2,1)=MP{1,i}(F100,2);F(2,2)=MP{1,i}(F100,1);F(2,3)=F(2,1)/F(2,2);
  F(3,1)=MP{1,i}(F150,2);F(3,2)=MP{1,i}(F150,1);F(3,3)=F(3,1)/F(3,2);
  F(4,1)=MP{1,i}(F200,2);F(4,2)=MP{1,i}(F200,1);F(4,3)=F(4,1)/F(4,2);
  F(5,1)=MP{1,i}(F250,2);F(5,2)=MP{1,i}(F250,1);F(5,3)=F(5,1)/F(5,2);
  F(6,1)=MP{1,i}(F300,2);F(6,2)=MP{1,i}(F300,1);F(6,3)=F(6,1)/F(6,2);
 kraft_weg(i,1)=F(6,1);kraft_weg(i,2)=F(6,2);kraft_weg(i,3)=F(6,3);%记录最大变形及刚度
 
     
 
    xlswrite(Fileaddress,{'Kraft(N)'},[strcat('Sheet',num2str(i))],'A1');
      xlswrite(Fileaddress,{'Weg(mm)'},[strcat('Sheet',num2str(i))],'B1');
      xlswrite(Fileaddress,{'Spez.Wert(N/mm)'},[strcat('Sheet',num2str(i))],'C1');
  xlswrite(Fileaddress,F,[strcat('Sheet',num2str(i))],'A2');
      sfilename=[Lstfhinten,'MP' num2str(i) '.jpg'];
          f=getframe(h);
           imwrite(f.cdata,sfilename);  
         waitbar(i/(length(MP)+1));
         close(h);
         Slidesn=Slides.Item(i);
 for m=1:6
     for n=1:3
         Slidesn.Shapes.Range.Item(3).Table.Cell(m+1,n).Shape.TextFrame.TextRange.Text=num2str(F(m,n),'%.0f');
         
         end
 end
 Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800);
        Slidesn.Shapes.Range.Item(2).Table.Cell(2,2).Shape.TextFrame.TextRange.Text=['MP',num2str(i),' bei RT'];
       Slidesn.Shapes.Range.Item(2).Table.Cell(1,2).Shape.TextFrame.TextRange.Text= get(handles.edit1,'String');
end

Presentation.Save
%system('taskkill/IM POWERPNT.exe');
p=find(kraft_weg(:,3)==min(kraft_weg(:,3)));
    MPwegmin=strcat('MP',num2str(p));
       Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));
     Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));
     Ezuobiao=strcat('E',num2str(SZ+1));Fzuobiao=strcat('F',num2str(SZ+1)); Gzuobiao=strcat('G',num2str(SZ+1));                               %生成写入EXCEL单元坐标
   xlswrite([Fileaddress1],{get(handles.edit1,'String')},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress1],{MPwegmin},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress1],kraft_weg(p,1),'Sheet1',[Czuobiao]);%写入C列对应变形
      xlswrite([Fileaddress1],kraft_weg(p,2),'Sheet1',[Dzuobiao]);%写入D列对应力
      xlswrite([Fileaddress1],kraft_weg(p,3),'Sheet1',[Ezuobiao]);%写入E列最小刚度
   xlswrite([Fileaddress1],{list{val1}},'Sheet1',[Fzuobiao]);%写入D列温度
   xlswrite([Fileaddress1],{date},'Sheet1',[Gzuobiao]);%写入E列时间
   
winopen(Lstfhinten)

end
waitbar(100);
close(t2);
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE=get(handles.edit1,'String');
TEST_NAME=['VW保险杠拉力试验 ',list{val1}];
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
msgbox('报告生成完毕');

% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)


% --------------------------------------------------------------------
function PPT2WORD_Callback(hObject, eventdata, handles)
[filename,pathname,fileindex]=uigetfile('*.ppt;*.pptx','选择ppt');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return
end
 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
filespec_user=fullfile(pathname,filename);
try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    


 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
 

  Fileaddress=fullfile(pathname,'result');
   Slides = Powerpoint.ActivePresentation.Slides; 
  Slides_number=Slides.count;
  t1=waitbar(0,'正在生成图片');
  for i=1:Slides_number
 Slides1=Slides.Item(i);
  Slides1.Export([Fileaddress,'\',num2str(i),'.bmp'],'bmp');
   pic = imread([Fileaddress,'\',num2str(i),'.bmp']);
   pic_1 = imcrop(pic,[26.25,100,990.5,555.5]);
   imwrite(pic_1,[Fileaddress,'\',num2str(i),'.bmp']);
   waitbar(i/(Slides_number));
  end
close(t1);
 try
       system('taskkill/IM POWERPNT.exe');
 end
       t2=waitbar(0,'正在生成Word报告');
 filespec_user=[Fileaddress,'\report.doc'];
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % 使word为可见；或set(Word, 'Visible', 1); 
%===打开word文件，如果路径下没有则创建一个空白文档打开========================
if exist(filespec_user,'file')
Document=Word.Documents.Open(filespec_user);
else
Document=Word.Documents.Add;
Document.SaveAs2(filespec_user);
end

Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='III. Einzelergebnis 具体结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落


InlineShapes=Document.InlineShapes;
He=180*1.26;
Wi=240*1.9;
for i=1:Slides_number
    sfilename1=[Fileaddress,'\',num2str(i),'.bmp'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
delete(sfilename1); 
Selection.Start = Selection.end;
Selection.TypeParagraph;
waitbar(i/Slides_number);
end
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
winopen([Fileaddress,'\report.doc']);
close(t2);


% --- Executes on button press in checkbox2.
function checkbox2_Callback(hObject, eventdata, handles)


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)
% hObject    handle to Fahrzeugcode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns Fahrzeugcode contents as cell array
%        contents{get(hObject,'Value')} returns selected item from Fahrzeugcode


% --- Executes during object creation, after setting all properties.
function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Fahrzeugcode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
