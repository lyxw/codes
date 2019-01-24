FROM ubuntu

RUN apt-get update && apt-get install -y ssh vim iproute2 net-tools iputils-ping wget curl && mkdir -p /run/sshd && mkdir -p /root/.ssh && mkdir -p /root/zookeeper/tmp

RUN ssh-keygen -t dsa -P "" -f /root/.ssh/id_rsa && cat /root/.ssh/id_rsa.pub >> /root/.ssh/authorized_keys && echo "/usr/sbin/sshd" >> /root/.bashrc && echo "RSAAuthentication yes" >> /etc/ssh/sshd_config && echo "PubkeyAuthentication yes" >> /etc/ssh/sshd_config && echo "AuthorizedKeysFile /root/.ssh/authorized_keys" >> /etc/ssh/sshd_config

COPY src/* /root/
WORKDIR /root

RUN tar zxf jdk-8u201-linux-x64.tar.gz && mv jdk1.8.0_201 /usr/jdk1.8.0_201 && tar zxf scala-2.12.8.tar.gz && mv scala-2.12.8 /usr/scala-2.12.8 && tar zxf zookeeper-3.4.13.tar.gz && mv zookeeper-3.4.13 /usr/zookeeper-3.4.13 && tar zxf hadoop-3.2.0.tar.gz && mv hadoop-3.2.0 /usr/hadoop-3.2.0 && tar zxf spark-2.4.0-bin-hadoop2.7.tgz && mv spark-2.4.0-bin-hadoop2.7 /usr/spark-2.4.0-bin-hadoop2.7 && rm -f jdk-8u201-linux-x64.tar.gz scala-2.12.8.tar.gz zookeeper-3.4.13.tar.gz hadoop-3.2.0.tar.gz spark-2.4.0-bin-hadoop2.7.tgz

RUN bash -c 'echo "export JAVA_HOME=/usr/jdk1.8.0_201" >> /root/.bashrc && echo "export SCALA_HOME=/usr/scala-2.12.8" >> /root/.bashrc && echo "export ZOOKEEPER_HOME=/usr/zookeeper-3.4.13" >> /root/.bashrc && echo "export HADOOP_HOME=/usr/hadoop-3.2.0" >> /root/.bashrc && echo "export SPARK_HOME=/usr/spark-2.4.0-bin-hadoop2.7" >> /root/.bashrc'

RUN mv -f zoo.cfg /usr/zookeeper-3.4.13/conf/zoo.cfg && mv -f core-site.xml /usr/hadoop-3.2.0/etc/hadoop/core-site.xml && mv -f hdfs-site.xml /usr/hadoop-3.2.0/etc/hadoop/hdfs-site.xml && mv -f yarn-site.xml /usr/hadoop-3.2.0/etc/hadoop/yarn-site.xml && mv -f workers /usr/hadoop-3.2.0/etc/hadoop/workers && mv -f spark-env.sh /usr/spark-2.4.0-bin-hadoop2.7/conf/spark-env.sh && mv -f slaves /usr/spark-2.4.0-bin-hadoop2.7/conf/slaves

RUN bash -c 'echo "export PATH=/usr/jdk1.8.0_201/bin:/usr/scala-2.12.8/bin:/usr/zookeeper-3.4.13/bin:/usr/hadoop-3.2.0/bin:/usr/hadoop-3.2.0/sbin:/usr/spark-2.4.0-bin-hadoop2.7/sbin:/usr/spark-2.4.0-bin-hadoop2.7/bin:$PATH" >> /root/.bashrc && echo 1 > /root/zookeeper/tmp/myid && echo "export JAVA_HOME=/usr/jdk1.8.0_201" >> /usr/hadoop-3.2.0/etc/hadoop/hadoop-env.sh'

CMD	["/bin/bash"]

# docker build -t ubuntu/spark:v1.1 .
