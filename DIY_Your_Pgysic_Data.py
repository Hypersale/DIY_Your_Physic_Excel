import pygame,os,openpyxl,keyboard,pyperclip,random
from math import sqrt
zoom_excel = 4
zoom_up = [1,2,5,8]
zoom_down = [1,0.5,0.2]
white = (255,255,255)
black = (0,0,0)
grey = (160,160,160)
light_grey = (200,200,200)
moving = False
class get_your_own_pyhsic_data(object):
    def __init__(self):
        pygame.init()
        self.size = (500,400) #width=500,height=300
        self.screen = pygame.display.set_mode(self.size,pygame.RESIZABLE)
        self.word_size = 22
        self.font = pygame.font.SysFont(pygame.font.get_default_font(),self.word_size)
        self.button_font = pygame.font.SysFont(pygame.font.get_default_font(),int(1.5*self.word_size))
        self.excel_min_y = 0
        self.excel_min_x = 0
        self.excel_block_x_zoom = 0
        self.excel_block_y_zoom = 0
        self.line_x_offset,self.line_y_offset = 0,0
        self.excel_height_lines = 10
        self.excel_width_lines = 10
        self.excel_width = 400
        self.excel_height = 200
        self.last = False
        self.dot_list = []
        self.random = 1
        self.round = 2
    def draw_frame(self):#default excel height=200,width=400
        self.window_real_width, self.window_real_height = pygame.display.get_surface().get_size()
        self.left_start = max(((100*self.window_real_width/1200)//10)*10,80)
        self.bottom_start = min(((200*self.window_real_height/200)//10)*10,((self.window_real_height//10)//1.5)*10)
        self.excel_height = self.bottom_start-20
        self.excel_width = self.window_real_width-self.left_start-20
        self.screen.fill(white)
        self.word_resize = min(self.window_real_width/500,self.window_real_height/300)
        self.font = pygame.font.SysFont(pygame.font.get_default_font(),round(self.word_size *self.word_resize))
        self.button_font = pygame.font.SysFont(pygame.font.get_default_font(),round(1.5 * self.word_size *self.word_resize))
        pygame.draw.line(self.screen,grey, (self.left_start, 20), (self.left_start,self.bottom_start), 2)
        pygame.draw.line(self.screen,grey, (self.left_start,self.bottom_start), (self.window_real_width-20, self.bottom_start), 2)
        return
    def draw_excel(self):
        self.zoom()
        self.draw_buttons()
        if moving:
            self.excel_move()
        if self.dot_list != []:
            self.draw_dots()
        for i in range(1,self.excel_height_lines+1): #draw excel lines from left to right
            x = self.zoom_excel+self.excel_block_x_zoom
            y = self.zoom_excel+self.excel_block_y_zoom
            if y >= 0:
                num = self.excel_min_y + (i-1)*self.excel_block_y
            else:
                num = round(self.excel_min_y + (i-1)*self.excel_block_y,abs(y)//3+1)
            self.screen.blit(self.font.render(str(num),True,black),(self.left_start-10-8*len(str(num)),self.line_y_offset+self.bottom_start-int(i*self.excel_height/self.excel_height_lines)-int(self.word_size*0.25)))
            pygame.draw.line(self.screen,light_grey,(self.left_start ,self.line_y_offset+self.bottom_start-int(i*self.excel_height/self.excel_height_lines)),(self.left_start+self.excel_width,self.line_y_offset+self.bottom_start-int(i*self.excel_height/self.excel_height_lines)))
        for i in range(self.excel_width_lines): #draw excel lines from up to down
            if x >= 0:
                num = self.excel_min_x + (i)*self.excel_block_x
            else:
                num = round(self.excel_min_x + (i)*self.excel_block_x,abs(x)//3+1)
            self.screen.blit(self.font.render(str(num),True,black),(self.line_x_offset-5*len(str(num))+self.left_start+int(i*self.excel_width/self.excel_width_lines),self.bottom_start+12))
            pygame.draw.line(self.screen,light_grey,(self.line_x_offset+self.left_start+int(i*self.excel_width/self.excel_width_lines),20),(self.line_x_offset+self.left_start+int(i*self.excel_width/self.excel_width_lines),self.bottom_start))
        return
    def zoom(self):
        self.zoom_excel = zoom_excel
        x = self.zoom_excel+self.excel_block_x_zoom
        y = self.zoom_excel+self.excel_block_y_zoom
        if x > 0:
            self.excel_block_x = (10**(x//4))*zoom_up[x%4]
            self.excel_min_x = int(self.excel_min_x-self.excel_min_x%self.excel_block_x)
        if x == 0:
            self.excel_block_x == 1
        if x < 0:
            self.excel_block_x = (0.1**(abs(x)//3))*zoom_down[abs(x)%3]
        if y > 0:
            self.excel_block_y = (10**(y//4))*zoom_up[y%4]
            self.excel_min_y = int(self.excel_min_x-self.excel_min_x%self.excel_block_y)
        elif y == 0:
            self.excel_block_y == 1
        elif y < 0:
            self.excel_block_y = (0.1**(abs(y)//3))*zoom_down[abs(y)%3]
    def excel_move(self):
        if self.last:
            self.last_x,self.last_y = self.excel_min_x,self.excel_min_y
            self.last = False
        new_x,new_y = pygame.mouse.get_pos()
        x_offset,y_offset = new_x-mouse_x,new_y-mouse_y
        self.line_x_offset,self.line_y_offset = x_offset%(self.excel_width/10),y_offset%(self.excel_height/10)
        self.excel_min_x = self.last_x - self.excel_block_x*(x_offset//(self.excel_width/10))
        self.excel_min_y = self.last_y + self.excel_block_y*(y_offset//(self.excel_height/10))
    def add_point(self):
        distance_list = []
        distance_1 = 10
        if self.dot_list != []:
            for dot in self.dot_list:
                x_pos = ((dot[0]-self.excel_min_x)  /  (self.excel_block_x/(self.excel_width/self.excel_width_lines)))  +  self.line_x_offset+self.left_start
                y_pos = (self.line_y_offset+self.bottom_start-int(self.excel_height/self.excel_height_lines))  -  ((dot[1]-self.excel_min_y)  /  (self.excel_block_y/(self.excel_height/self.excel_height_lines)))
                distance_list.append(sqrt((mouse_x-x_pos)*(mouse_x-x_pos)+(mouse_y-y_pos)*(mouse_y-y_pos)))
            distance_1 = distance_list[0]
            distance_2 = 0
            for i in range(1,len(distance_list)):
                print(distance_1,distance_2)
                if distance_list[i] < distance_1:
                    distance_1 = distance_list[i]
                    distance_2 = i
        if distance_1 <= 8:
            self.dot_list.pop(distance_2)# right click on the dot will remove the dot
        else:
            distance_x_to_min = mouse_x - (self.line_x_offset+self.left_start)
            distance_y_to_min = (self.line_y_offset+self.bottom_start-int(self.excel_height/self.excel_height_lines)) - mouse_y
            num_x_per_pixel = self.excel_block_x/(self.excel_width/self.excel_width_lines)
            num_y_per_pixel = self.excel_block_y/(self.excel_height/self.excel_height_lines)
            num_1 = distance_x_to_min*num_x_per_pixel + self.excel_min_x
            num_2 = distance_y_to_min*num_y_per_pixel + self.excel_min_y
            self.dot_list.append([num_1,num_2])
            self.sort_dotslist()
            print(num_1,num_2)
    def draw_dots(self):
        first = True
        for dot in self.dot_list:
            x_pos = ((dot[0]-self.excel_min_x)  /  (self.excel_block_x/(self.excel_width/self.excel_width_lines)))  +  self.line_x_offset+self.left_start
            y_pos = (self.line_y_offset+self.bottom_start-int(self.excel_height/self.excel_height_lines))  -  ((dot[1]-self.excel_min_y)  /  (self.excel_block_y/(self.excel_height/self.excel_height_lines)))
            if not first:
                pygame.draw.line(self.screen,black,(last_x_pos,last_y_pos),(x_pos,y_pos))
            pygame.draw.circle(self.screen, black, (x_pos,y_pos), 5, 5)
            first = False
            last_x_pos = x_pos
            last_y_pos = y_pos
    def sort_dotslist(self):
        for i in range(len(self.dot_list)-1):    
            for j in range(len(self.dot_list)-i-1):  
                if self.dot_list[j][0] > self.dot_list[j+1][0]:
                    p = self.dot_list[j]
                    self.dot_list.pop(j)
                    self.dot_list.insert(j+1,p)
    def draw_buttons(self):
        word_size = self.word_size*self.word_resize
        self.screen.blit(self.button_font.render("+",True,black),(self.window_real_width-self.excel_width/10,self.bottom_start+word_size))#zoom in the right bottom corner
        self.screen.blit(self.button_font.render("-",True,black),(int(self.window_real_width-self.excel_width/10-0.8*word_size),self.bottom_start+word_size))
        self.screen.blit(self.button_font.render("+",True,black),(int(self.left_start-10-1.5*word_size),20))#zoom in the left top corner
        self.screen.blit(self.button_font.render("-",True,black),(int(self.left_start-10-1.5*word_size),20+word_size))
        self.screen.blit(self.button_font.render("Out Put Excel",True,black),(int(self.left_start-10-1.5*word_size),int(self.bottom_start+1.5*word_size)))
        self.screen.blit(self.button_font.render("Start Pasting",True,black),(int(self.left_start-10+self.excel_width/5+2*word_size),int(self.bottom_start+1.5*word_size)))
        self.screen.blit(self.button_font.render("Stop",True,black),(int(self.left_start-10+self.excel_width//1.8+3*word_size),int(self.bottom_start+1.5*word_size)))
        self.screen.blit(self.button_font.render(("Random: "+str(self.random)+"%"),True,black),(int(self.left_start-10-1.5*word_size),int(self.bottom_start+2.5*word_size)))
        self.screen.blit(self.button_font.render(("Decimal point: "+str(self.round)),True,black),(int(self.left_start-10+5.5*word_size),int(self.bottom_start+2.5*word_size)))
    def out_put_excel(self):
        self.randomize_dotlist()
        self.current_path = os.getcwd()
        current_workbook = openpyxl.Workbook()
        worksheet = current_workbook.active
        for i in range(len(self.dot_list)):
            worksheet.cell(column=i+1,row=1).value = round(self.dot_list[i][0],self.round)
            worksheet.cell(column=i+1,row=2).value = round(self.dot_list[i][1],self.round)
        current_workbook.save(self.current_path+"\\OutPut_Excel.xlsx")
    def randomize_dotlist(self):
        if len(self.dot_list) > 1 and self.random != 0:
            distance_list = []
            for i in range(1,len(self.dot_list)):
                distance_list.append(self.dot_list[i-1][1]-self.dot_list[i][1])
            for i in range(1,len(self.dot_list)):
                self.dot_list[i][1] += (distance_list[i-1]*random.uniform(-1,1)*self.random)/100
fwork = get_your_own_pyhsic_data()
run = True
global paste_time
paste_time = 0
start = False
button_round = False
button_random = False
def pasting_data_counting(len1):
    global paste_time
    if paste_time < 2*len1:
        list1 = ["x","y"]
        print("Pasted dot_",(paste_time)%len1+1,",",list1[paste_time//len1],":",round(fwork.dot_list[paste_time%len1][paste_time//len1],fwork.round))
        pyperclip.copy(str(round(fwork.dot_list[paste_time%len1][paste_time//len1],fwork.round)))
        paste_time += 1
    else:
        pyperclip.copy("End")
while run:
    fwork.draw_frame()
    fwork.draw_excel()
    # textSurface = get_your_own_pyhsic_data.font.render("CSDNM114", True, black)
    # get_your_own_pyhsic_data.screen.blit(textSurface, (100,100))
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            run = False
        if event.type == pygame.MOUSEBUTTONDOWN:
            mouse_x,mouse_y = pygame.mouse.get_pos()
            word_size = fwork.word_size*fwork.word_resize
            if fwork.left_start < mouse_x < fwork.window_real_width-20 and 20 < mouse_y < fwork.bottom_start:
                if event.button == 4:#滚轮上
                    zoom_excel += 1
                if event.button == 5:#滚轮下
                    zoom_excel -= 1
                if event.button == 1:#按下左键
                    moving = True
                    fwork.last = True
                if event.button == 3:#按下右键
                    fwork.add_point()
            elif int(fwork.left_start-10-1.5*word_size) < mouse_x < int(fwork.left_start-10+5*word_size) and int(fwork.bottom_start+1.5*word_size) < mouse_y < int(fwork.bottom_start+2.5*word_size):
                if len(fwork.dot_list):
                    fwork.out_put_excel()
                    print("Already Out Put Excel On"+fwork.current_path+"\\OutPut_Excel.xlsx")
            elif int(fwork.left_start-10+fwork.excel_width/5+2*word_size) < mouse_x < int(fwork.left_start-10+fwork.excel_width/5+8.5*word_size) and int(fwork.bottom_start+1.5*word_size) < mouse_y < int(fwork.bottom_start+2.5*word_size):
                if len(fwork.dot_list):
                    start = True
                    paste_time = 0
                    fwork.randomize_dotlist()
                    pyperclip.copy(str(fwork.dot_list[paste_time%len(fwork.dot_list)][paste_time//len(fwork.dot_list)]))
                    keyboard.add_hotkey('ctrl+v', pasting_data_counting, args=(len(fwork.dot_list),))
                    print("Now start pasting data.")
            elif fwork.window_real_width-fwork.excel_width/10 < mouse_x < fwork.window_real_width-fwork.excel_width/10+0.6*word_size and fwork.bottom_start+word_size < mouse_y < fwork.bottom_start+1.6*word_size:
                fwork.excel_block_x_zoom += 1
            elif int(fwork.window_real_width-fwork.excel_width/10-0.8*word_size) < mouse_x < int(fwork.window_real_width-fwork.excel_width/10-0.2*word_size) and fwork.bottom_start+word_size < mouse_y < fwork.bottom_start+1.6*word_size:
                fwork.excel_block_x_zoom -= 1
            elif int(fwork.left_start-10-1.5*word_size) < mouse_x < int(fwork.left_start-10-0.8*word_size) and 20 < mouse_y < 20+word_size:
                fwork.excel_block_y_zoom += 1
            elif int(fwork.left_start-10-1.5*word_size) < mouse_x < int(fwork.left_start-10-0.8*word_size) and 20+word_size < mouse_y < 20+2*word_size:
                fwork.excel_block_y_zoom -= 1
            elif int(fwork.left_start-10-1.5*word_size) < mouse_x < int(fwork.left_start-10+4.5*word_size) and int(fwork.bottom_start+2.5*word_size) < mouse_y < int(fwork.bottom_start+3.5*word_size):
                button_random = True
                fwork.random = 0
            elif int(fwork.left_start-10+5.5*word_size) < mouse_x < int(fwork.left_start-10+13.5*word_size) and int(fwork.bottom_start+2.5*word_size) < mouse_y < int(fwork.bottom_start+3.5*word_size):
                button_round = True
                fwork.round = 0
        if start:
            if int(fwork.left_start-10+fwork.excel_width//1.8+3*word_size) < mouse_x < int(fwork.left_start-10+fwork.excel_width//1.8+5*word_size) and int(fwork.bottom_start+1.5*word_size) < mouse_y < int(fwork.bottom_start+2.5*word_size):
                start = False
                print("Stop pasting data")
                keyboard.remove_hotkey('ctrl+v')
            if paste_time >= 2*len(fwork.dot_list):
                start = False
                print("Stop pasting data")
                keyboard.remove_hotkey('ctrl+v')
        if button_round == True:
            if event.type == pygame.KEYDOWN:
                if 48 <= event.key <= 57:
                    fwork.round = 10*fwork.round + int(event.unicode)
                if event.key == 13 or event.unicode == '\r':
                    button_round = False
        if button_random == True:
            if event.type == pygame.KEYDOWN:
                if 48 <= event.key <= 57:
                    fwork.random = 10*fwork.random + int(event.unicode)
                if event.key == 13 or event.unicode == '\r':
                    button_random = False
        if event.type == pygame.MOUSEBUTTONUP and event.button == 1:
            moving = False
    pygame.display.update()