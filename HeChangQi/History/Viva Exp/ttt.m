%v=zeros(5000000,20);
b=0;
while true
    b=b+1;
a=input('','s');
a=[a '.wav'];
t=wavread(a);
v(1:size(t,1),b)=t(:,1);
end
