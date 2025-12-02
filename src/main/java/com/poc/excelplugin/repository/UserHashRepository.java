package com.poc.excelplugin.repository;

import com.poc.excelplugin.entity.UserFileHash;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface UserHashRepository extends JpaRepository<UserFileHash, Long> {

    // Custom query to validate if the file record exists
    boolean existsByUserIdAndHashKey(String userId, String hashKey);
}